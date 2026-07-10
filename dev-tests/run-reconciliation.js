// Headless payroll reconciliation runner.
//
// Drives the REAL payroll-reconciliation.html pipeline (processReconciliation →
// journal → audit PDF) in Node using the vendored XLSX + jsPDF libraries, with
// input files parsed exactly like the browser (sheet_to_json header:1).
//
// Usage:  node run-reconciliation.js <config.json>
// Config: { client, period, transactionListing, remuneration, leaveThis,
//           leaveLast ("APRIL_SNAPSHOT" uses the saved prior leaveThisSnapshot),
//           contacts, outDir }
//
// No side effects on master data: fetch and localStorage are stubbed, so the
// OneDrive history/data JSONs are never written. The snapshot the app builds is
// captured to <outDir>/<period>-snapshot.json instead.
const fs = require('fs');
const path = require('path');

const cfg = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));

// ---- real libraries (vendored) ----
// Load UMD builds with module/exports/define shadowed so they attach to the
// global object like they would in a browser.
global.window = global;
global.self = global;
function loadUMD(file) {
  const src = fs.readFileSync(path.join(__dirname, file), 'utf8');
  new Function('module', 'exports', 'define', src).call(global, undefined, undefined, undefined);
}
loadUMD('vendor-xlsx.full.min.js');
loadUMD('vendor-jspdf.umd.min.js');
const XLSXlib = global.XLSX;
// jsPDF's browser save() can't run in Node — write the buffer to outDir instead.
const { jsPDF } = global.jspdf;
jsPDF.API.save = function (filename) {
  const out = path.join(cfg.outDir, filename);
  fs.writeFileSync(out, Buffer.from(this.output('arraybuffer')));
  console.log('PDF written:', out);
};

// ---- DOM stub ----
const elements = {};
function makeEl(id) {
  return {
    id, value: '', innerHTML: '', style: {}, textContent: '', checked: false, files: [],
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    appendChild(){}, remove(){}, addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    scrollIntoView(){}, querySelector(){ return null; }, querySelectorAll(){ return []; },
  };
}
const outputDiv = makeEl('output');
function parseTable() {
  if (!outputDiv.innerHTML.includes('<table')) return null;
  const rowsHtml = [...outputDiv.innerHTML.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)];
  const bodyRows = rowsHtml.filter(r => !r[1].includes('<th'));
  return {
    querySelectorAll: () => bodyRows.map(r => {
      const cells = [...r[1].matchAll(/<td[^>]*>([\s\S]*?)<\/td>/g)].map(c => ({
        innerText: c[1].replace(/<[^>]*>/g, '').trim(), colSpan: /colspan/.test(c[0]) ? 2 : 1
      }));
      return { classList: { contains: () => /class="[^"]*highlight/.test(r[0]) }, querySelectorAll: () => cells };
    })
  };
}
global.document = {
  getElementById(id) { if (id === 'output') return outputDiv; if (!elements[id]) elements[id] = makeEl(id); return elements[id]; },
  createElement(tag) { return makeEl(tag); },
  querySelector(sel) { return sel === '#output table' ? parseTable() : null; },
  querySelectorAll() { return []; },
  body: { appendChild(){} },
  addEventListener(){},
};
global.addEventListener = () => {}; global.removeEventListener = () => {};
global.setInterval = () => 0; global.setTimeout = (fn) => 0;
global.alert = (m) => console.log('[alert]', String(m).split('\n')[0]);
global.confirm = (m) => { console.log('[confirm→OK]', String(m).split('\n')[0]); return true; };
global.localStorage = { getItem(){ return null; }, setItem(){}, removeItem(){} };
global.sessionStorage = global.localStorage;
global.location = { hash: '', origin: '', pathname: '', search: '' };
global.fetch = () => Promise.resolve({ ok: false, json: () => ({}) });
global.navigator = { userAgent: 'node-headless' };
global.FileReader = function(){};

// ---- parse inputs exactly like the app's readExcel ----
function readExcelFile(p) {
  const wb = XLSXlib.read(fs.readFileSync(p), { type: 'buffer' });
  return XLSXlib.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null });
}

// ---- client record from master data JSON (read-only) ----
const DATA = 'C:/Users/f869f/OneDrive - LF Advisory Pty Ltd/LF Advisory Workpapers/lf-advisory-payroll-data.json';
const clientData = JSON.parse(fs.readFileSync(DATA, 'utf8'))[cfg.client];
if (!clientData) throw new Error('Client not found in payroll data: ' + cfg.client);

global.__inputs = {
  transactionData: readExcelFile(cfg.transactionListing),
  remunerationData: readExcelFile(cfg.remuneration),
  leaveThisData: readExcelFile(cfg.leaveThis),
  leaveLastData: cfg.leaveLast === 'PRIOR_SNAPSHOT'
    ? (clientData.lastReconciliation && clientData.lastReconciliation.leaveThisSnapshot)
    : readExcelFile(cfg.leaveLast),
  contactsData: readExcelFile(cfg.contacts),
  clientData,
  cfg,
};
if (!global.__inputs.leaveLastData) throw new Error('No last-month leave data available');

// ---- load app code + drive it inside the same eval scope ----
const html = fs.readFileSync(path.join(__dirname, '..', 'payroll-reconciliation.html'), 'utf8');
const scripts = [...html.matchAll(/<script(?![^>]*src)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);

const driver = `
// inject parsed files (globals are let-bound in this scope)
transactionData  = global.__inputs.transactionData;
remunerationData = global.__inputs.remunerationData;
leaveThisData    = global.__inputs.leaveThisData;
leaveLastData    = global.__inputs.leaveLastData;
contactsData     = global.__inputs.contactsData;

// client settings (mirrors loadClientSettings, which normally reads localStorage)
currentClientData = JSON.parse(JSON.stringify(global.__inputs.clientData));
departments = (currentClientData.departments || []).map(d => ({...d, glCodes: d.glCodes || {}}));

document.getElementById('clientName').value = global.__inputs.cfg.client;
document.getElementById('period').value = global.__inputs.cfg.period;
document.getElementById('workcover').value = String(currentClientData.workcover || 0);
document.getElementById('payrollFrequency').value = currentClientData.payrollFrequency || 'monthly';
document.getElementById('accrualMethod').value = currentClientData.accrualMethod || 'standard';
document.getElementById('autoCalculatePayrollTax').checked = currentClientData.autoCalculatePayrollTax !== false;
document.getElementById('groupedEmployer').checked = false;
document.getElementById('groupWages').value = '0';
document.getElementById('nextPayDate').value = '';

// CoA inputs from saved chart of accounts (defaults apply when absent)
const __coa = currentClientData.chartOfAccounts || {};
for (const [k, v] of Object.entries(__coa)) {
  document.getElementById('coa_' + k).value = v;
}

// opening balances = prior rec closing balances (as loadClientSettings does)
if (currentClientData.lastReconciliation && currentClientData.lastReconciliation.closingBalances) {
  for (const [deptId, bal] of Object.entries(currentClientData.lastReconciliation.closingBalances)) {
    document.getElementById('opening_' + deptId).value = String(bal);
  }
}

processReconciliation();

global.__run = {
  lslData: window.lastLSLData,
  journalEntries: journalEntries,
  snapshot: null,
};
`;

try { eval(scripts.join('\n;\n') + '\n;\n' + driver); }
catch (e) { console.error('RUN FAILED:', e.message, '\n', e.stack.split('\n').slice(0, 5).join('\n')); process.exit(1); }

const run = global.__run;
console.log('\n=== RUN COMPLETE:', cfg.client, '—', cfg.period, '===');
console.log('journal lines:', run.journalEntries.length);
run.journalEntries.forEach(e => {
  if (e.debit > 0 || e.credit > 0)
    console.log(`  ${String(e.account).padEnd(6)} ${e.name.padEnd(55)} DR ${e.debit.toFixed(2).padStart(10)}  CR ${e.credit.toFixed(2).padStart(10)}`);
});
if (run.lslData) {
  console.log('\nLSL provision:');
  (run.lslData.employeeProvisions || []).forEach(e =>
    console.log(`  ${e.name.padEnd(18)} ${e.state}${e.stateDefaulted ? '(defaulted)' : ''} ${e.vestingYears}y svc=${e.yearsOfService.toFixed(2)} prov=$${e.provision.toFixed(2)} ${e.isCurrent ? 'CURRENT' : 'NON-CURRENT'}`));
  console.log(`  Current $${run.lslData.currentLiability.toFixed(2)} | Non-current $${run.lslData.nonCurrentLiability.toFixed(2)} | Movement $${run.lslData.movement.toFixed(2)}`);
} else {
  console.log('\nWARNING: no LSL data produced');
}

// capture the snapshot the app would have saved (history write is stubbed off)
try {
  const snapRaw = JSON.parse(JSON.stringify({
    period: cfg.period,
    lslData: global.__run.lslData,
    journalEntries: run.journalEntries,
  }));
  fs.writeFileSync(path.join(cfg.outDir, cfg.period.replace(/ /g, '-') + '-headless-run.json'), JSON.stringify(snapRaw, null, 2));
} catch (e) { console.warn('snapshot capture failed:', e.message); }

// generate the audit PDF with the real jsPDF (function declarations from the
// first eval are bound in this module scope)
try {
  generateAuditReport(); // uses window.lastLSLData + on-screen journal table
} catch (e) {
  console.error('PDF FAILED:', e.message);
  process.exit(1);
}
