// Parity harness for report-formatter.html: replays the page's real inline
// script outside the browser (DOM stubbed, ExcelJS from node_modules) so the
// portal's workbook can be diffed against the Python writer's with
// tools/compare_workbooks.py in Projects/lfa-report-formatter.
//
//   node test-report-formatter.js <raw-xero-export.xlsx> <out.xlsx>
//
// Needs exceljs resolvable — set NODE_PATH to a folder that has it, e.g.
//   NODE_PATH=/path/with/node_modules node test-report-formatter.js in.xlsx out.xlsx
const fs = require('fs');
const path = require('path');

const [, , IN, OUT] = process.argv;
if (!IN || !OUT) {
  console.error('usage: node test-report-formatter.js <raw-export.xlsx> <out.xlsx>');
  process.exit(2);
}

function makeEl(id) {
  return {
    id, value: '', innerHTML: '', textContent: '', checked: false, style: {},
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    appendChild(){}, remove(){}, addEventListener(){}, setAttribute(){},
    getAttribute(){ return null; }, scrollIntoView(){},
    querySelector(){ return null; }, querySelectorAll(){ return []; },
    files: [],
  };
}
const elements = {};
global.document = {
  getElementById(id) { if (!elements[id]) elements[id] = makeEl(id); return elements[id]; },
  createElement() { return makeEl('el'); },
  querySelector() { return null; }, querySelectorAll() { return []; },
  body: { appendChild(){} }, addEventListener(){},
};
global.window = global;
global.addEventListener = () => {};
global.alert = () => {};
global.localStorage = { getItem(){ return null; }, setItem(){}, removeItem(){} };
global.sessionStorage = global.localStorage;
global.location = { hash: '', origin: '', pathname: '', search: '' };
global.fetch = () => Promise.resolve({ ok: false, json: () => ({}) });
global.navigator = { userAgent: 'node' };
global.URL = { createObjectURL(){ return ''; }, revokeObjectURL(){} };
global.Blob = class {};
global.ExcelJS = require('exceljs');

const html = fs.readFileSync(path.join(__dirname, '..', 'report-formatter.html'), 'utf8');
// the page has one <script> block after the markup
const m = html.match(/<script>\s*'use strict';([\s\S]*?)<\/script>/);
if (!m) { console.error('could not find the inline script'); process.exit(2); }
eval(m[1] + '\n;global.__api = {parseManagementPack, transformPack, checkPack, clientConfigFor, buildWorkbook};');

(async () => {
  const api = global.__api;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(IN);
  const { reports, skipped } = api.parseManagementPack(wb);
  const kinds = Object.keys(reports);
  if (!kinds.length) { console.error('no recognised report sheets'); process.exit(1); }
  console.log(`parsed: ${kinds.join(', ')}${skipped.length ? `  (excluded: ${skipped.join(', ')})` : ''}`);
  const org = (Object.values(reports)[0] || {}).orgName || '';
  const pack = api.transformPack(reports, api.clientConfigFor(org));
  const fails = api.checkPack(pack);
  for (const f of fails) console.log('  ' + f);
  const out = api.buildWorkbook(pack);
  await out.xlsx.writeFile(OUT);
  console.log(`wrote ${OUT} (${pack.sheets.length} sheets, client: ${pack.clientName})`);
  for (const s of pack.sheets) console.log(`  ${s.sheetName}: ${s.periodicity}`);
})().catch(e => { console.error(e); process.exit(1); });
