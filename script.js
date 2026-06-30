const fileInput = document.getElementById('fileInput');
const dropzone = document.getElementById('dropzone');
const downloadBtn = document.getElementById('downloadBtn');
const copyXmlBtn = document.getElementById('copyXmlBtn');
const outputName = document.getElementById('outputName');
const summary = document.getElementById('summary');
const warningsEl = document.getElementById('warnings');
const previewTable = document.getElementById('previewTable');
const ignoredTable = document.getElementById('ignoredTable');
const xmlPreview = document.getElementById('xmlPreview');
const statRows = document.getElementById('statRows');
const statExported = document.getElementById('statExported');
const statIgnored = document.getElementById('statIgnored');
const statWarnings = document.getElementById('statWarnings');

let generatedXml = '';
let allRows = [];
let exportRows = [];
let ignoredRows = [];
let checks = [];

const XML_NS = 'http://schema.csn.se/Studeranderapport';
const DATE_FIELDS = ['fromDatum', 'tomDatum', 'resultatDatum', 'andringsDatum', 'avbrottsDatum'];
const FIELD_ORDER = ['fromDatum', 'tomDatum', 'omfattning', 'resultat', 'resultatDatum', 'andringsDatum', 'avbrottsDatum', 'uppdragsutbildning', 'resultatAvbrott'];
const PREVIEW_COLS = ['_rowNumber', '_status', 'skolId', 'personnummer', 'fromDatum', 'tomDatum', 'omfattning', 'resultat', 'resultatDatum', 'avbrottsDatum', 'resultatAvbrott'];
const REQUIRED_EXPORT_FIELDS = ['personnummer', 'fromDatum', 'tomDatum', 'omfattning'];

function normalizeHeader(header) {
  return String(header || '').trim().replace(/^ns\d*:/i, '').replace(/^.*:/, '');
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function excelSerialDateToIso(value) {
  const parsed = XLSX.SSF.parse_date_code(value);
  if (!parsed) return String(value);
  return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
}

function formatDate(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date && !isNaN(value)) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') return excelSerialDateToIso(value);
  const s = String(value).trim();
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return s;
}

function normalizePersonnummer(value) {
  if (value == null || value === '') return '';
  return String(value).trim().replace(/\D/g, '');
}

function normalizeDecimal(value) {
  if (value == null || value === '') return '';
  return String(value).trim().replace(',', '.');
}

function normalizeValue(key, value) {
  if (value == null || value === '') return '';
  if (DATE_FIELDS.includes(key)) return formatDate(value);
  if (key === 'personnummer') return normalizePersonnummer(value);
  if (key === 'omfattning') return normalizeDecimal(value);
  if (['resultat', 'resultatAvbrott', 'uppdragsutbildning'].includes(key)) return String(value).trim().toUpperCase();
  return String(value).trim();
}

function isEmptyRow(row) {
  return Object.entries(row).every(([key, value]) => key === '_rowNumber' || value === '');
}

function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
  return rows.map((row, idx) => {
    const out = { _rowNumber: idx + 2 };
    Object.entries(row).forEach(([header, value]) => {
      const key = normalizeHeader(header);
      if (key) out[key] = normalizeValue(key, value);
    });
    return out;
  }).filter(row => !isEmptyRow(row));
}

function rowExportEligibility(row) {
  const missing = REQUIRED_EXPORT_FIELDS.filter(field => !row[field]);
  if (missing.length) {
    return { exportable: false, reason: `Saknar obligatoriskt fält: ${missing.join(', ')}` };
  }
  return { exportable: true, reason: 'Exporteras' };
}

function annotateRows(rows) {
  return rows.map(row => {
    const eligibility = rowExportEligibility(row);
    return { ...row, _exportable: eligibility.exportable, _reason: eligibility.reason };
  });
}

function validateRows(rows) {
  const result = [];
  const schoolIds = new Set(rows.map(r => r.skolId).filter(Boolean));
  if (schoolIds.size > 1) result.push({ text: `Flera skolId hittades i rader som exporteras: ${[...schoolIds].join(', ')}. Kontrollera att detta är avsiktligt.`, type: 'warn' });

  rows.forEach(row => {
    const prefix = `Rad ${row._rowNumber}`;
    if (!row.skolId) result.push({ text: `${prefix}: saknar skolId.`, type: 'bad' });
    if (!/^\d{12}$/.test(row.personnummer || '')) result.push({ text: `${prefix}: personnummer ska vara 12 siffror efter konvertering.`, type: 'bad' });
    ['fromDatum','tomDatum'].forEach(k => {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(row[k] || '')) result.push({ text: `${prefix}: ${k} saknas eller har fel datumformat.`, type: 'bad' });
    });
    if (!row.omfattning) result.push({ text: `${prefix}: saknar omfattning.`, type: 'bad' });
    if (row.omfattning && !/^\d+(\.\d+)?$/.test(row.omfattning)) result.push({ text: `${prefix}: omfattning ska vara ett tal, exempelvis 45 eller 44.5.`, type: 'bad' });
    if (row.resultat && !['J','N'].includes(row.resultat)) result.push({ text: `${prefix}: resultat ska vara J eller N.`, type: 'bad' });
    if (row.resultatAvbrott && !['J','N'].includes(row.resultatAvbrott)) result.push({ text: `${prefix}: resultatAvbrott ska vara J eller N.`, type: 'bad' });
    if (row.resultat && !row.resultatDatum) result.push({ text: `${prefix}: resultat finns men resultatDatum saknas.`, type: 'warn' });
    if (row.avbrottsDatum && !row.resultatAvbrott) result.push({ text: `${prefix}: avbrottsDatum finns men resultatAvbrott saknas.`, type: 'warn' });
    if (row.resultat && row.avbrottsDatum) result.push({ text: `${prefix}: både resultat och avbrottsDatum finns. Kontrollera raden.`, type: 'warn' });
  });
  return result;
}

function buildXml(rows) {
  const schoolId = rows.find(r => r.skolId)?.skolId || '';
  const lines = [];
  lines.push('<?xml version="1.0" encoding="utf-16"?>');
  lines.push(`<studeranderapport xmlns="${XML_NS}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="${XML_NS} Studeranderapport.xsd">`);
  lines.push('\t<skola>');
  lines.push(`\t\t<skolId>${escapeXml(schoolId)}</skolId>`);
  rows.forEach(row => {
    lines.push('\t\t<studerande>');
    lines.push(`\t\t\t<personnummer>${escapeXml(row.personnummer || '')}</personnummer>`);
    lines.push('\t\t\t<studieperiod>');
    FIELD_ORDER.forEach(key => {
      const value = row[key];
      if (value !== undefined && value !== null && value !== '') {
        lines.push(`\t\t\t\t<${key}>${escapeXml(value)}</${key}>`);
      }
    });
    lines.push('\t\t\t</studieperiod>');
    lines.push('\t\t</studerande>');
  });
  lines.push('\t</skola>');
  lines.push('</studeranderapport>');
  return lines.join('\r\n');
}

function renderList(el, items, emptyText) {
  el.innerHTML = '';
  if (!items.length) {
    const li = document.createElement('li');
    li.textContent = emptyText;
    li.className = 'ok';
    el.appendChild(li);
    return;
  }
  items.forEach(item => {
    const li = document.createElement('li');
    li.textContent = item.text || item;
    li.className = item.type || '';
    el.appendChild(li);
  });
}

function renderSummary() {
  const students = new Set(exportRows.map(r => r.personnummer).filter(Boolean));
  statRows.textContent = allRows.length;
  statExported.textContent = exportRows.length;
  statIgnored.textContent = ignoredRows.length;
  statWarnings.textContent = checks.length;

  renderList(summary, [
    { text: `${allRows.length} ifyllda rader hittades i Excel-filen.`, type: allRows.length ? 'ok' : 'bad' },
    { text: `${exportRows.length} rader kommer exporteras till XML.`, type: exportRows.length ? 'ok' : 'bad' },
    { text: `${ignoredRows.length} rader ignoreras.`, type: ignoredRows.length ? 'warn' : 'ok' },
    { text: `${students.size} unika studerande exporteras.`, type: 'ok' },
    { text: checks.length ? `${checks.length} saker behöver kontrolleras innan uppladdning till CSN.` : 'Inga uppenbara fel hittades i exporten.', type: checks.length ? 'warn' : 'ok' }
  ], 'Ingen sammanfattning ännu.');
}

function renderWarnings(warningList) {
  const ignoredChecks = ignoredRows.map(row => ({ text: `Rad ${row._rowNumber} ignoreras: ${row._reason}.`, type: 'warn' }));
  renderList(warningsEl, [...warningList, ...ignoredChecks], 'Inga uppenbara fel hittades.');
}

function maskPersonnummer(value) {
  const s = String(value || '');
  if (s.length < 6) return s;
  return `${s.slice(0, 6)}******`;
}

function renderStatusCell(td, row) {
  const span = document.createElement('span');
  span.className = row._exportable ? 'status-pill status-ok' : 'status-pill status-warn';
  span.textContent = row._exportable ? 'Exporteras' : 'Ignoreras';
  td.appendChild(span);
}

function renderTable(table, rows) {
  table.innerHTML = '';
  if (!rows.length) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.textContent = 'Inga rader att visa.';
    tr.appendChild(td);
    table.appendChild(tr);
    return;
  }
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  PREVIEW_COLS.forEach(c => {
    const th = document.createElement('th');
    th.textContent = c === '_rowNumber' ? 'Excel-rad' : c === '_status' ? 'Status' : c;
    trh.appendChild(th);
  });
  const reasonTh = document.createElement('th');
  reasonTh.textContent = 'Kommentar';
  trh.appendChild(reasonTh);
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  rows.forEach(row => {
    const tr = document.createElement('tr');
    if (!row._exportable) tr.className = 'warn';
    PREVIEW_COLS.forEach(c => {
      const td = document.createElement('td');
      if (c === '_status') renderStatusCell(td, row);
      else if (c === 'personnummer') td.textContent = maskPersonnummer(row[c]);
      else td.textContent = row[c] || '';
      tr.appendChild(td);
    });
    const tdReason = document.createElement('td');
    tdReason.textContent = row._reason || '';
    tr.appendChild(tdReason);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

function utf16Blob(text) {
  const buffer = new ArrayBuffer(2 + text.length * 2);
  const view = new DataView(buffer);
  view.setUint16(0, 0xFEFF, true);
  for (let i = 0; i < text.length; i++) view.setUint16(2 + i * 2, text.charCodeAt(i), true);
  return new Blob([buffer], { type: 'application/xml;charset=utf-16' });
}

async function handleFile(file) {
  try {
    const buffer = await file.arrayBuffer();
    allRows = annotateRows(parseWorkbook(buffer));
    exportRows = allRows.filter(row => row._exportable);
    ignoredRows = allRows.filter(row => !row._exportable);
    checks = validateRows(exportRows);
    generatedXml = exportRows.length ? buildXml(exportRows) : '';

    renderSummary();
    renderWarnings(checks);
    renderTable(previewTable, exportRows);
    renderTable(ignoredTable, ignoredRows);
    xmlPreview.textContent = generatedXml ? generatedXml.slice(0, 12000) + (generatedXml.length > 12000 ? '\n\n… XML:en är längre, men hela laddas ner.' : '') : 'Ingen XML skapad eftersom inga exportbara rader hittades.';
    downloadBtn.disabled = exportRows.length === 0;
    copyXmlBtn.disabled = exportRows.length === 0;

    if (!outputName.value.trim() || outputName.value === 'csn-rapport.xml') {
      const base = file.name.replace(/\.[^.]+$/, '').replace(/\s+/g, '-').toLowerCase();
      outputName.value = `${base}.xml`;
    }
  } catch (err) {
    console.error(err);
    renderList(warningsEl, [{ text: 'Filen kunde inte läsas. Kontrollera att det är en Excel-fil enligt CSN-mallen.', type: 'bad' }], '');
  }
}

fileInput.addEventListener('change', event => {
  const file = event.target.files?.[0];
  if (file) handleFile(file);
});

['dragenter','dragover'].forEach(evt => dropzone.addEventListener(evt, e => {
  e.preventDefault();
  dropzone.classList.add('dragover');
}));
['dragleave','drop'].forEach(evt => dropzone.addEventListener(evt, e => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
}));
dropzone.addEventListener('drop', e => {
  const file = e.dataTransfer.files?.[0];
  if (file) handleFile(file);
});

downloadBtn.addEventListener('click', () => {
  const blob = utf16Blob(generatedXml);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = outputName.value.trim().endsWith('.xml') ? outputName.value.trim() : `${outputName.value.trim() || 'csn-rapport'}.xml`;
  document.body.appendChild(a);
  a.click();
  URL.revokeObjectURL(a.href);
  a.remove();
});

copyXmlBtn.addEventListener('click', async () => {
  await navigator.clipboard.writeText(generatedXml);
  copyXmlBtn.textContent = 'Kopierad!';
  setTimeout(() => copyXmlBtn.textContent = 'Kopiera XML', 1200);
});
