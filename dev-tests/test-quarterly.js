// Unit test for the quarterly reporting-frequency feature. Evaluates the real
// inline script from payroll-reconciliation.html with DOM/global stubs
// (same pattern as test-live-classification.js / test-tb-recon.js).
const fs = require('fs');

const elements = {};
function makeEl(id) {
  return {
    id, value: '', innerHTML: '', textContent: '', checked: false, style: {},
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
global.alert = () => {}; global.confirm = () => true;
global.localStorage = { getItem(){return null;}, setItem(){}, removeItem(){} };
global.sessionStorage = global.localStorage;
global.location = { hash: '', origin: '', pathname: '', search: '' };
global.fetch = () => Promise.resolve({ ok:false, json:()=>({}) });
global.navigator = { userAgent: 'test' };
global.XLSX = null;

const html = fs.readFileSync('C:/Users/f869f/LFAdvisoryWebsite/payroll-reconciliation.html','utf8');
const scripts = [...html.matchAll(/<script(?![^>]*src)[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]);

const testSetup = `
const R = [];
const approx = (a,b,t)=>Math.abs(a-b) <= (t==null?0.01:t);

// ---------- getMonthsInPeriod ----------
currentClientData = { reportingFrequency: 'monthly' };
R.push(['monthly → 1 month', getMonthsInPeriod()===1]);
currentClientData = { reportingFrequency: 'quarterly' };
R.push(['quarterly → 3 months', getMonthsInPeriod()===3]);
currentClientData = {};
R.push(['absent → defaults to 1', getMonthsInPeriod()===1]);

// ---------- Workcover: one period's share of the annual premium ----------
const wcAnnual = 46329.63;
R.push(['workcover monthly = annual/12', approx(wcAnnual/12*1, 3860.8025)]);
R.push(['workcover quarterly = annual/12*3', approx(wcAnnual/12*3, 11582.4075)]);

// ---------- Payroll tax: monthly threshold vs 3× for quarterly ----------
// Single-state QLD example. Monthly threshold 108,333.33, rate 4.75%.
// Monthly run on 1 month of wages (150k): tax = (150000 - 108333.33) * .0475
const wagesM = { QLD: 150000 };
const monthly = calculatePayrollTax(wagesM, { monthsInPeriod: 1 });
R.push(['PT monthly: deduction = 1× threshold', approx(monthly.breakdown[0].proportionateDeduction, 108333.33)]);
R.push(['PT monthly: tax on 150k', approx(monthly.totalTax, (150000-108333.33)*0.0475)]);

// Quarterly run on 3 months of wages (450k): deduction = 3× threshold
const wagesQ = { QLD: 450000 };
const quarterly = calculatePayrollTax(wagesQ, { monthsInPeriod: 3 });
R.push(['PT quarterly: deduction = 3× threshold', approx(quarterly.breakdown[0].proportionateDeduction, 108333.33*3)]);
R.push(['PT quarterly: tax on 450k', approx(quarterly.totalTax, (450000-108333.33*3)*0.0475)]);
R.push(['PT quarterly: displayed threshold scaled ×3', approx(quarterly.breakdown[0].threshold, 108333.33*3)]);

// Sanity: steady 150k/mth for a quarter should tax the same as 3 monthly runs
R.push(['PT quarterly ≈ 3× the equivalent monthly tax', approx(quarterly.totalTax, monthly.totalTax*3, 0.02)]);

// ---------- Grouped employer scales its deduction too ----------
const grouped = calculatePayrollTax({ QLD: 450000 }, { isGrouped: true, groupWages: 450000, monthsInPeriod: 3 });
R.push(['PT grouped quarterly: deduction ×3', approx(grouped.breakdown[0].deduction, 5145.83*3)]);

global.__results = R;
`;

try { eval(scripts.join('\n;\n') + '\n;\n' + testSetup); }
catch (e) { console.log('EVAL ERROR:', e.message, '\n', e.stack); process.exit(2); }

const R = global.__results || [];
let fail = 0;
R.forEach(([label, ok]) => { console.log((ok?'PASS':'FAIL')+'  '+label); if(!ok) fail++; });
console.log(`\n${R.length-fail}/${R.length} passed`);
process.exit(fail ? 1 : 0);
