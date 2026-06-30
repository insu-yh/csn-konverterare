const REQUIRED = ['personnummer','fromDatum','tomDatum','omfattning'];
const XML_FIELDS = ['fromDatum','tomDatum','omfattning','resultat','resultatDatum','andringsDatum','avbrottsDatum','uppdragsutbildning','resultatAvbrott'];
let latestXml = '';

const $ = id => document.getElementById(id);
const els = {
  fileInput: $('fileInput'), dropzone: $('dropzone'), outputName: $('outputName'), downloadBtn: $('downloadBtn'), copyXmlBtn: $('copyXmlBtn'),
  statRows: $('statRows'), statExported: $('statExported'), statIgnored: $('statIgnored'), statWarnings: $('statWarnings'),
  summary: $('summary'), warnings: $('warnings'), previewTable: $('previewTable'), ignoredTable: $('ignoredTable'), xmlPreview: $('xmlPreview'), debugPanel: $('debugPanel')
};

els.fileInput.addEventListener('change', e => handleFile(e.target.files[0]));
['dragenter','dragover'].forEach(ev => els.dropzone.addEventListener(ev, e => { e.preventDefault(); els.dropzone.classList.add('drag'); }));
['dragleave','drop'].forEach(ev => els.dropzone.addEventListener(ev, e => { e.preventDefault(); els.dropzone.classList.remove('drag'); }));
els.dropzone.addEventListener('drop', e => handleFile(e.dataTransfer.files[0]));
els.downloadBtn.addEventListener('click', downloadXml);
els.copyXmlBtn.addEventListener('click', async () => { await navigator.clipboard.writeText(latestXml); els.copyXmlBtn.textContent = 'Kopierad!'; setTimeout(()=>els.copyXmlBtn.textContent='Kopiera XML', 1200); });

function handleFile(file){
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try{
      const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array', cellDates:true});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {defval:null, raw:false});
      const normalized = rows.map((row, i) => normalizeRow(row, i + 2)).filter(r => r.hasAnyData);
      const checked = normalized.map(validateRow);
      const exportRows = checked.filter(r => r.exportable);
      const ignoredRows = checked.filter(r => !r.exportable);
      latestXml = buildXml(exportRows);
      render(checked, exportRows, ignoredRows, file.name);
    }catch(err){
      alert('Kunde inte läsa filen: ' + err.message);
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

function stripHeader(key){ return String(key || '').replace(/^.*:/,'').trim(); }
function normalizeRow(row, excelRow){
  const clean = { excelRow, raw: row };
  for(const [key, value] of Object.entries(row)) clean[stripHeader(key)] = normalizeValue(stripHeader(key), value);
  clean.personnummer = normalizePersonnummer(clean.personnummer);
  clean.hasAnyData = Object.entries(clean).some(([k,v]) => !['excelRow','raw','hasAnyData'].includes(k) && v !== null && v !== undefined && String(v).trim() !== '');
  return clean;
}
function normalizeValue(field, value){
  if(value === null || value === undefined) return '';
  if(value instanceof Date && !isNaN(value)) return toIsoDate(value);
  const text = String(value).trim();
  if(!text) return '';
  if(/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  if(/^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}$/.test(text)){
    const parts = text.split(/[/. -]/).filter(Boolean).map(Number);
    let [d,m,y] = parts; if(y < 100) y += 2000;
    return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
  }
  return text;
}
function toIsoDate(d){ return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
function normalizePersonnummer(value){
  const s = String(value || '').trim(); if(!s) return '';
  return s.replace(/\D/g,'');
}
function maskPnr(p){
  const s = String(p || '');
  if(s.length < 8) return s ? '***' : '';
  return `${s.slice(0,8)}-****`;
}
function validDate(s){ return /^\d{4}-\d{2}-\d{2}$/.test(String(s||'')); }

function validateRow(row){
  const checks = [];
  const missing = REQUIRED.filter(f => !row[f]);
  for(const f of REQUIRED) checks.push({field:f, value:row[f], status:row[f]?'ok':'bad', message:row[f]?'Finns':'Saknas'});
  const exportable = missing.length === 0;
  if(row.personnummer) checks.push({field:'personnummer format', value:maskPnr(row.personnummer), status:/^\d{12}$/.test(row.personnummer)?'ok':'warn', message:/^\d{12}$/.test(row.personnummer)?'12 siffror':'Borde vara 12 siffror efter rensning'});
  for(const f of ['fromDatum','tomDatum','resultatDatum','andringsDatum','avbrottsDatum']) if(row[f]) checks.push({field:f+' format', value:row[f], status:validDate(row[f])?'ok':'warn', message:validDate(row[f])?'YYYY-MM-DD':'Kontrollera datumformat'});

  if(row.resultat && !row.resultatDatum) checks.push({field:'resultatDatum', value:'', status:'warn', message:'Resultat finns men resultatDatum saknas'});
  if(row.avbrottsDatum && !row.resultatAvbrott) checks.push({field:'resultatAvbrott', value:'', status:'warn', message:'Avbrottsdatum finns men resultatAvbrott saknas'});
  if(row.resultat && row.avbrottsDatum) checks.push({field:'resultat/avbrott', value:`${row.resultat} / ${row.avbrottsDatum}`, status:'warn', message:'Raden har både resultat och avbrott'});
  if(row.omfattning && isNaN(Number(String(row.omfattning).replace(',','.')))) checks.push({field:'omfattning', value:row.omfattning, status:'warn', message:'Omfattning är inte ett tal'});
  const warnings = checks.filter(c => c.status === 'warn');
  const errors = checks.filter(c => c.status === 'bad');
  const type = row.avbrottsDatum ? 'Avbrott' : row.resultat ? 'Resultat' : 'Studieperiod';
  return {...row, exportable, warnings, errors, checks, type, ignoreReason: missing.length ? `Saknar ${missing.join(', ')}` : ''};
}

function buildXml(rows){
  const skolId = rows.find(r => r.skolId)?.skolId || '35620';
  const lines = ['<?xml version="1.0" encoding="utf-16"?>','<studeranderapport xmlns="http://schema.csn.se/Studeranderapport" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://schema.csn.se/Studeranderapport Studeranderapport.xsd">','\t<skola>',`\t\t<skolId>${escapeXml(skolId)}</skolId>`];
  rows.forEach(r => {
    lines.push('\t\t<studerande>', `\t\t\t<personnummer>${escapeXml(r.personnummer)}</personnummer>`, '\t\t\t<studieperiod>');
    XML_FIELDS.forEach(f => { if(r[f] !== '' && r[f] !== null && r[f] !== undefined) lines.push(`\t\t\t\t<${f}>${escapeXml(formatField(f,r[f]))}</${f}>`); });
    lines.push('\t\t\t</studieperiod>', '\t\t</studerande>');
  });
  lines.push('\t</skola>','</studeranderapport>');
  return lines.join('\n');
}
function formatField(f,v){ if(f === 'omfattning') return String(v).replace(',','.'); return String(v); }
function escapeXml(s){ return String(s).replace(/[<>&'"]/g, ch => ({'<':'&lt;','>':'&gt;','&':'&amp;',"'":'&apos;','"':'&quot;'}[ch])); }

function render(all, exported, ignored, filename){
  els.statRows.textContent = all.length; els.statExported.textContent = exported.length; els.statIgnored.textContent = ignored.length;
  const warnCount = all.reduce((sum,r)=>sum+r.warnings.length,0); els.statWarnings.textContent = warnCount;
  els.downloadBtn.disabled = exported.length === 0; els.copyXmlBtn.disabled = exported.length === 0;
  els.summary.innerHTML = [
    li('ok', `Filen lästes in: ${filename}`), li(exported.length?'ok':'bad', `${exported.length} rader kommer exporteras till XML`), li(ignored.length?'warn':'ok', `${ignored.length} rader ignoreras`), li(warnCount?'warn':'ok', `${warnCount} varningar hittades`)
  ].join('');
  const warnItems = all.flatMap(r => r.warnings.map(w => li('warn', `Rad ${r.excelRow}: ${w.message}`)));
  els.warnings.innerHTML = warnItems.length ? warnItems.join('') : li('ok','Inga varningar hittades.');
  renderTable(els.previewTable, exported, ['excelRow','personnummer','fromDatum','tomDatum','omfattning','type','resultat','resultatDatum','avbrottsDatum','resultatAvbrott'], true);
  renderTable(els.ignoredTable, ignored, ['excelRow','personnummer','fromDatum','tomDatum','omfattning','ignoreReason'], true);
  renderDebug(all);
  els.xmlPreview.textContent = latestXml || 'Ingen XML skapad ännu.';
}
function li(cls,text){ return `<li class="${cls}">${escapeHtml(text)}</li>`; }
function renderTable(table, rows, cols, mask){
  if(!rows.length){ table.innerHTML = '<tbody><tr><td>Inga rader.</td></tr></tbody>'; return; }
  table.innerHTML = `<thead><tr>${cols.map(c=>`<th>${label(c)}</th>`).join('')}</tr></thead><tbody>` + rows.map(r => `<tr>${cols.map(c => `<td>${cell(c,r,mask)}</td>`).join('')}</tr>`).join('') + '</tbody>';
}
function cell(c,r,mask){
  if(c === 'personnummer') return escapeHtml(mask ? maskPnr(r[c]) : r[c]);
  if(c === 'type') return `<span class="pill ${r.warnings.length?'warn':'ok'}">${escapeHtml(r.type)}</span>`;
  if(c === 'ignoreReason') return `<span class="pill bad">${escapeHtml(r.ignoreReason)}</span>`;
  return escapeHtml(r[c] ?? '');
}
function renderDebug(rows){
  if(!rows.length){ els.debugPanel.className='debug-list empty'; els.debugPanel.textContent='Inga rader att visa.'; return; }
  els.debugPanel.className='debug-list';
  els.debugPanel.innerHTML = rows.map(r => {
    const status = !r.exportable ? 'bad' : r.warnings.length ? 'warn' : 'ok';
    const statusText = !r.exportable ? 'Ignoreras' : r.warnings.length ? 'Exporteras med varning' : 'Exporteras';
    return `<details class="debug-item"><summary><span><span class="debug-title">Rad ${r.excelRow} · ${maskPnr(r.personnummer)} · ${escapeHtml(r.type)}</span><br><span class="debug-meta">${escapeHtml(statusText)}</span></span><span class="pill ${status}">${statusIcon(status)} ${escapeHtml(statusText)}</span></summary><div class="debug-body">${r.checks.map(ch => `<div class="debug-check"><strong>${escapeHtml(ch.field)}</strong><code>${escapeHtml(ch.value || '—')}</code><span class="${ch.status}">${statusIcon(ch.status)} ${escapeHtml(ch.message)}</span></div>`).join('')}</div></details>`;
  }).join('');
}
function statusIcon(s){ return s === 'ok' ? '✓' : s === 'warn' ? '⚠' : '✗'; }
function label(c){ return ({excelRow:'Rad',personnummer:'Personnummer',fromDatum:'Från',tomDatum:'Till',omfattning:'Omfattning',type:'Typ',resultat:'Resultat',resultatDatum:'Resultatdatum',avbrottsDatum:'Avbrottsdatum',resultatAvbrott:'Resultat avbrott',ignoreReason:'Orsak'}[c] || c); }
function escapeHtml(s){ return String(s ?? '').replace(/[<>&'"]/g, ch => ({'<':'&lt;','>':'&gt;','&':'&amp;',"'":'&#39;','"':'&quot;'}[ch])); }
function encodeUtf16Le(str){
  const buf = new ArrayBuffer(2 + str.length * 2); const view = new DataView(buf); view.setUint16(0, 0xFEFF, true);
  for(let i=0;i<str.length;i++) view.setUint16(2 + i*2, str.charCodeAt(i), true); return buf;
}
function downloadXml(){
  const name = (els.outputName.value || 'csn-rapport.xml').replace(/[^a-zA-Z0-9åäöÅÄÖ_.-]/g,'_');
  const blob = new Blob([encodeUtf16Le(latestXml)], {type:'application/xml;charset=utf-16'});
  const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = name.endsWith('.xml') ? name : name + '.xml'; a.click(); URL.revokeObjectURL(url);
}
