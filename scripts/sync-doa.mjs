import fs from 'node:fs/promises';
import * as XLSX from 'xlsx';

const OUTPUT_PATH = process.env.DOA_OUTPUT_PATH || new URL('../doa-data.json', import.meta.url);
const sharedLink = process.env.ONEDRIVE_DOA_URL?.trim();

if (!sharedLink) {
  throw new Error('Falta el secreto ONEDRIVE_DOA_URL con el enlace compartido de OneDrive.');
}

const norm = (value = '') => String(value ?? '')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .trim()
  .toUpperCase();

const idValue = (value) => String(value ?? '')
  .trim()
  .replace(/\.0+$/, '')
  .toUpperCase();

function dateRank(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.getTime();
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Date.UTC(1899, 11, 30) + Math.round(value) * 86400000;
  }
  const text = String(value ?? '').trim();
  let match = text.match(/\b([0-3]?\d)[-\/]([0-1]?\d)[-\/](20\d{2}|\d{2})\b/);
  if (match) {
    const year = Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
    return Date.UTC(year, Number(match[2]) - 1, Number(match[1]));
  }
  return 0;
}

function headerRowIndex(rows) {
  let best = { index: 0, score: -1 };
  rows.slice(0, 12).forEach((row, index) => {
    const values = row.map(norm);
    const score = values.filter((value) => value === 'UPC').length * 20
      + values.filter((value) => value === 'CLU' || value.includes('PART NUMBER')).length * 16
      + values.filter((value) => value.includes('DESCRIP')).length * 12
      + values.filter((value) => value.includes('NUMERO DE SERIE')).length * 10
      + values.filter(Boolean).length;
    if (score > best.score) best = { index, score };
  });
  return best.index;
}

async function downloadWorkbook() {
  const downloadUrl = new URL(sharedLink);
  downloadUrl.searchParams.set('download', '1');
  const response = await fetch(downloadUrl, {
    redirect: 'follow',
    headers: {
      accept: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel;q=0.9,*/*;q=0.5',
      'user-agent': 'Inventory-Lens-DOA-Sync/1.0',
    },
  });

  if (!response.ok) {
    throw new Error(`OneDrive respondió ${response.status}. Revisa que el enlace permita acceso anónimo y descarga.`);
  }

  const bytes = new Uint8Array(await response.arrayBuffer());
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
  if (!isZip) {
    throw new Error('OneDrive devolvió una página web en lugar del XLSX. El vínculo puede requerir inicio de sesión o bloquear la descarga.');
  }
  return bytes;
}

function extractRecords(workbookBytes) {
  const workbook = XLSX.read(workbookBytes, { type: 'array', cellDates: true });
  const identifiable = [];
  const incomplete = [];

  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: null,
      raw: true,
    });
    if (!rows.length) continue;

    const headerRow = headerRowIndex(rows);
    const headers = (rows[headerRow] || []).map(norm);
    const index = {
      upc: headers.findIndex((header) => header === 'UPC' || header.includes('UPC')),
      clu: headers.findIndex((header) => header === 'CLU' || header.includes('PART NUMBER')),
      serial: headers.findIndex((header) => header.includes('NUMERO DE SERIE')),
      changeDate: headers.findIndex((header) => header.includes('FECHA EN QUE SE HIZO EL CAMBIO DOA')),
      description: headers.findIndex((header) => header.includes('DESCRIP')),
    };

    rows.slice(headerRow + 1).forEach((row, offset) => {
      const record = {
        store: sheetName.trim(),
        row: headerRow + offset + 2,
        upc: index.upc >= 0 ? idValue(row[index.upc]) : '',
        clu: index.clu >= 0 ? idValue(row[index.clu]) : '',
        serial: index.serial >= 0 ? idValue(row[index.serial]) : '',
        changedAt: index.changeDate >= 0 ? dateRank(row[index.changeDate]) : 0,
        description: index.description >= 0 ? String(row[index.description] ?? '').trim() : '',
      };
      if (!(record.upc || record.clu || record.serial || record.description)) return;
      if (!(record.upc || record.clu)) incomplete.push(record);
      else identifiable.push(record);
    });
  }

  const bySerial = new Map();
  const withoutSerial = [];
  const duplicateSerials = new Set();
  identifiable.forEach((record) => {
    if (!record.serial) {
      withoutSerial.push(record);
      return;
    }
    const current = bySerial.get(record.serial);
    if (!current || record.changedAt >= current.changedAt) {
      if (current) duplicateSerials.add(record.serial);
      bySerial.set(record.serial, record);
    } else {
      duplicateSerials.add(record.serial);
    }
  });

  const accepted = [...bySerial.values(), ...withoutSerial];
  const grouped = new Map();
  accepted.forEach((record) => {
    const key = [norm(record.store), record.upc, record.clu].join('|');
    const aggregate = grouped.get(key) || {
      store: record.store,
      upc: record.upc,
      clu: record.clu,
      count: 0,
    };
    aggregate.count += 1;
    grouped.set(key, aggregate);
  });

  const records = [...grouped.values()].sort((a, b) =>
    a.store.localeCompare(b.store, 'es')
    || a.upc.localeCompare(b.upc)
    || a.clu.localeCompare(b.clu));

  return {
    records,
    quality: {
      rawRows: identifiable.length + incomplete.length,
      acceptedDevices: accepted.length,
      incompleteRows: incomplete.length,
      duplicateSerialsResolved: duplicateSerials.size,
    },
  };
}

async function existingSnapshot() {
  try {
    return JSON.parse(await fs.readFile(OUTPUT_PATH, 'utf8'));
  } catch {
    return null;
  }
}

const workbookBytes = await downloadWorkbook();
const extracted = extractRecords(workbookBytes);
const previous = await existingSnapshot();
const sameData = previous
  && JSON.stringify(previous.records) === JSON.stringify(extracted.records)
  && JSON.stringify(previous.quality) === JSON.stringify(extracted.quality);

const snapshot = {
  schemaVersion: 1,
  source: 'OneDrive · Listado de equipos DOA.xlsx',
  updatedAt: sameData && previous.updatedAt ? previous.updatedAt : new Date().toISOString(),
  records: extracted.records,
  quality: extracted.quality,
};

await fs.writeFile(OUTPUT_PATH, `${JSON.stringify(snapshot, null, 2)}\n`);
console.log(`DOA sincronizado: ${snapshot.quality.acceptedDevices} equipos, ${snapshot.records.length} grupos, ${snapshot.quality.incompleteRows} filas incompletas.`);
