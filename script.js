const fileInput = document.getElementById('fileInput');
const dropzone = document.getElementById('dropzone');
const downloadBtn = document.getElementById('downloadBtn');
const outputName = document.getElementById('outputName');
const summary = document.getElementById('summary');
const warningsEl = document.getElementById('warnings');
const previewTable = document.getElementById('previewTable');

let generatedXml = '';
let parsedRows = [];
let warnings = [];

const XML_NS = 'http://schema.csn.se/Studeranderapport';
const FIELD_ORDER = [
  'fromDatum',
  'tomDatum',
  'omfattning',
  'resultat',
  'resultatDatum',
  'andringsDatum',
  'avbrottsDatum',
  'uppdragsutbildning',
  'resultatAvbrott'
];

function normalizeHeader(header) {
  return String(header || '').trim().replace(/^ns\d*:/i, '');
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function formatDate(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date && !isNaN(value)) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return String(value);
    const mm = String(parsed.m).padStart(2, '0');
    const dd = String(parsed.d).padStart(2, '0');
    return `${parsed.y}-${mm}-${dd}`;
  }
  const s = String(value).trim();
  const m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return s;
}

function normalizePersonnummer(value) {
  if (value == null || value === '') return '';
  return String(value).trim().replace(/\D/g, '');
}

function normalizeValue(key, value) {
  if (value == null || value === '') return '';
  if (['fromDatum','tomDatum','resultatDatum','andringsDatum','avbrottsDatum'].includes(key)) return formatDate(value);
  if (key === 'personnummer') return normalizePersonnummer(value);
  if (key === 'resultat' || key === 'resultatAvbrott' || key === 'uppdragsutbildning') return String(value).trim().toUpperCase();
  return String(value).trim();
}

function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
  return rows.map((row, idx) => {
    const out = { _rowNumber: idx + 2 };
    Object.entries(row).forEach(([header, value]) => {
      const key = normalizeHeader(header);
      out[key] = normalizeValue(key, value);
    });
    return out;
  }).filter(row => Object.entries(row).some(([key, value]) => key !== '_rowNumber' && value !== ''));
}

function validateRows(rows) {
  const result = [];
  rows.forEach(row => {
    const prefix = `Rad ${row._rowNumber}`;
    if (!row.skolId) result.push(`${prefix}: saknar skolId.`);
    if (!/^\d{12}$/.test(row.personnummer || '')) result.push(`${prefix}: personnummer ska vara 12 siffror efter konvertering.`);
    ['fromDatum','tomDatum'].forEach(k => {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(row[k] || '')) result.push(`${prefix}: ${k} saknas eller har fel datumformat.`);
    });
    if (!row.omfattning) result.push(`${prefix}: saknar omfattning.`);
    if (row.resultat && !['J','N'].includes(row.resultat)) result.push(`${prefix}: resultat ska vara J eller N.`);
    if (row.resultatAvbrott && !['J','N'].includes(row.resultatAvbrott)) result.push(`${prefix}: resultatAvbrott ska vara J eller N.`);
    if (row.resultat && !row.resultatDatum) result.push(`${prefix}: resultat finns men resultatDatum saknas.`);
    if (row.avbrottsDatum && !row.resultatAvbrott) result.push(`${prefix}: avbrottsDatum finns men resultatAvbrott saknas.`);
    if (row.resultat && row.avbrottsDatum) result.push(`${prefix}: både resultat och avbrottsDatum finns. Kontrollera att det verkligen ska vara så.`);
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

function renderSummary(rows, warnings) {
  summary.innerHTML = '';
  const add = (text, cls = '') => {
    const li = document.createElement('li');
    li.textContent = text;
    if (cls) li.className = cls;
    summary.appendChild(li);
  };
  add(`${rows.length} rader hittades.`, 'ok');
  add(`${new Set(rows.map(r => r.skolId).filter(Boolean)).size} skolId hittades.`);
  add(`${warnings.length} varningar/fel hittades.`, warnings.length ? 'warning' : 'ok');
}

function renderWarnings(warnings) {
  warningsEl.innerHTML = '';
  if (!warnings.length) {
    const li = document.createElement('li');
    li.textContent = 'Inga uppenbara fel hittades.';
    li.className = 'ok';
    warningsEl.appendChild(li);
    return;
  }
  warnings.forEach(w => {
    const li = document.createElement('li');
    li.textContent = w;
    li.className = 'warning';
    warningsEl.appendChild(li);
  });
}

function renderPreview(rows) {
  const cols = ['skolId','personnummer','fromDatum','tomDatum','omfattning','resultat','resultatDatum','avbrottsDatum','resultatAvbrott'];
  previewTable.innerHTML = '';
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  cols.forEach(c => { const th = document.createElement('th'); th.textContent = c; trh.appendChild(th); });
  thead.appendChild(trh);
  previewTable.appendChild(thead);
  const tbody = document.createElement('tbody');
  rows.forEach(row => {
    const tr = document.createElement('tr');
    cols.forEach(c => { const td = document.createElement('td'); td.textContent = row[c] || ''; tr.appendChild(td); });
    tbody.appendChild(tr);
  });
  previewTable.appendChild(tbody);
}

async function handleFile(file) {
  const buffer = await file.arrayBuffer();
  parsedRows = parseWorkbook(buffer);
  warnings = validateRows(parsedRows);
  generatedXml = buildXml(parsedRows);
  renderSummary(parsedRows, warnings);
  renderWarnings(warnings);
  renderPreview(parsedRows);
  downloadBtn.disabled = parsedRows.length === 0;
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
  const blob = new Blob([generatedXml], { type: 'application/xml;charset=utf-16' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = outputName.value.trim() || 'csn-rapport.xml';
  document.body.appendChild(a);
  a.click();
  URL.revokeObjectURL(a.href);
  a.remove();
});
