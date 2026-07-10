// Live-engine test: run the edited calculateLSLProvision with fabricated June 2026
// inputs for Zummo and check the current/non-current classification.
const fs = require('fs');

// ---- minimal DOM stub (period input matters here) ----
const elements = {};
function makeEl(id) {
  return {
    id, value: '', innerHTML: '', style: {}, textContent: '', checked: false,
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    appendChild(){}, remove(){}, addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    scrollIntoView(){}, querySelector(){ return null; }, querySelectorAll(){ return []; },
  };
}
global.document = {
  getElementById(id) { if (!elements[id]) elements[id] = makeEl(id); return elements[id]; },
  createElement(tag) { return makeEl(tag); },
  querySelector() { return null; }, querySelectorAll() { return []; },
  body: { appendChild(){} }, addEventListener(){},
};
global.window = global;
global.addEventListener = () => {}; global.removeEventListener = () => {};
global.setInterval = () => 0; global.setTimeout = () => 0;
global.alert = () => {}; global.confirm = () => false;
global.localStorage = { getItem(){return null;}, setItem(){}, removeItem(){} };
global.sessionStorage = global.localStorage;
global.location = { hash: '', origin: '', pathname: '', search: '' };
global.fetch = () => Promise.resolve({ ok:false, json:()=>({}) });
global.navigator = { userAgent: 'test' };
global.jspdf = { jsPDF: null };
global.XLSX = null;

const html = fs.readFileSync('C:/Users/f869f/LFAdvisoryWebsite/payroll-reconciliation.html','utf8');
const scripts = [...html.matchAll(/<script(?![^>]*src)[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]);

// The app's globals are let-bound inside the eval scope, so the test setup must
// run in that same scope — append it to the eval'd source.
const h = JSON.parse(fs.readFileSync('C:/Users/f869f/OneDrive - LF Advisory Pty Ltd/LF Advisory Workpapers/lf-advisory-payroll-history.json','utf8'));
const snap = h['Zummo Juicers Pty Ltd']['April 2026'];
global.__snap = snap;

const testSetup = `
document.getElementById('period').value = 'June 2026';

// Leave files: reuse April's real leave file for both months (movement ≈ 0; classification is what we test)
leaveThisData = global.__snap.leaveThisSnapshot;
leaveLastData = global.__snap.leaveThisSnapshot;

// Remuneration file: header + [name, startDate DD/MM/YYYY, term, ..., salary@9]
remunerationData = [
  ['Remuneration'],
  ['Employee','Start Date','Termination Date','x','x','x','x','x','x','Annual Salary'],
  ["Adrian O'Connor",'16/03/2020','', '','','','','','', '130000'],
  ['Marcelle Horn','31/12/2019','', '','','','','','', '135000'],
  ['Shelly Brown','06/10/2020','', '','','','','','', '93500'],
];

// Contacts file: rows from index 4, employee col 0, state col 6.
// Deliberately OMIT Shelly Brown to prove the override + defaulted flag paths.
contactsData = [
  ['Contacts'],[],[],[],
  ["Adrian O'Connor",'','','','','','VIC'],
  ['Marcelle Horn','','','','','','QLD'],
];

// Client settings incl. the new state override for Shelly Brown, as now stored in the data JSON
currentClientData = { lslSettings: { enabled: true, stateOverrides: { 'Shelly Brown': 'NSW' } } };
departments = [{ id: 1, name: 'Overhead employees', xeroGroup: 'None', catchAll: true, glCodes: {} }];

global.__result = calculateLSLProvision();
`;

try { eval(scripts.join('\n;\n') + '\n;\n' + testSetup); } catch (e) { console.log('EVAL ERROR:', e.message); }
const result = global.__result;
console.log('\n================ RESULT (period end 30 June 2026) ================');
result.employeeProvisions.forEach(e => {
  console.log(`${e.name.padEnd(18)} state=${e.state}${e.stateDefaulted?'(defaulted)':''} vest=${e.vestingYears}y years=${e.yearsOfService.toFixed(2)} provision=$${e.provision.toFixed(2)} -> ${e.isCurrent ? 'CURRENT' : 'NON-CURRENT'}`);
});
console.log(`currentLiability    = $${result.currentLiability.toFixed(2)}`);
console.log(`nonCurrentLiability = $${result.nonCurrentLiability.toFixed(2)}`);
console.log(`currentMovement     = $${result.currentMovement.toFixed(2)} | nonCurrentMovement = $${result.nonCurrentMovement.toFixed(2)}`);

// Assertions
const adrian = result.employeeProvisions.find(e => /Adrian/.test(e.name));
const shelly = result.employeeProvisions.find(e => /Shelly/.test(e.name));
const marcelle = result.employeeProvisions.find(e => /Marcelle/.test(e.name));
const checks = [
  ['Adrian is VIC/7yr', adrian.state === 'VIC' && adrian.vestingYears === 7],
  ['Adrian years ~6.29 at 30 Jun 2026 (period end, not today)', Math.abs(adrian.yearsOfService - 6.29) < 0.02],
  ['Adrian classified CURRENT (vests 16 Mar 2027 ≤ 12 mths)', adrian.isCurrent === true],
  ['Shelly state = NSW via override, not defaulted', shelly.state === 'NSW' && !shelly.stateDefaulted],
  ['Shelly NON-CURRENT (5.7 yrs of 10)', shelly.isCurrent === false],
  ['Marcelle NON-CURRENT (6.5 yrs of 10)', marcelle.isCurrent === false],
  ['currentLiability equals Adrian provision', Math.abs(result.currentLiability - adrian.provision) < 0.01],
];
let fail = 0;
checks.forEach(([label, ok]) => { console.log((ok ? 'PASS' : 'FAIL') + '  ' + label); if (!ok) fail++; });
process.exit(fail ? 1 : 0);
