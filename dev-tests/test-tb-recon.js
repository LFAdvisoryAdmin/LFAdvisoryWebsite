// Unit test for the Trial Balance carried-forward feature, evaluating the real
// inline script from payroll-reconciliation.html with DOM/global stubs
// (same pattern as dev-tests/test-live-classification.js).
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
  querySelector(sel) {
    const m = /name="tbAlloc_([^"]+)"/.exec(sel || '');
    if (m) return { value: (global.__radioChoice && global.__radioChoice[m[1]]) || 'prop' };
    return null;
  },
  querySelectorAll() { return []; },
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
function setOpening(id, v){ document.getElementById('opening_'+id).value = String(v); }
function getOpening(id){ return parseFloat(document.getElementById('opening_'+id).value)||0; }

// continueReconciliation would run the full rec — stub it out for the apply tests.
continueReconciliation = function(){ global.__continued = true; };

// ---------- Test 1: parseTrialBalance (Xero single-Account-column format) ----------
{
  const rows = [
    ['Zummo Juicers Pty Ltd'], ['Trial Balance'], ['End of Mar 2024'], [],
    ['Account','Debit','Credit'],
    ['21381 - Annual Leave Provision', null, 50000],
    ['21382 - LSL Provision Current', null, 12000],
    ['24800 - LSL Provision Non-Current', null, 8000],
    ['21420 - Wages Payable', null, 3000],
    ['11195 - Workcover Prepaid', 4000, null],
    ['Total - Liabilities', null, 73000],
  ];
  const map = parseTrialBalance(rows);
  R.push(['Xero: AL liability = 50000', approx(tbAccountBalance(map,'21381','liability'),50000)]);
  R.push(['Xero: workcover asset = 4000', approx(tbAccountBalance(map,'11195','asset'),4000)]);
  R.push(['Xero: LSL current = 12000', approx(tbAccountBalance(map,'21382','liability'),12000)]);
  R.push(['Xero: Total- row skipped (no code "Total")', tbAccountBalance(map,'Total','liability')===null]);
}

// ---------- Test 2: parseTrialBalance (MYOB separate-columns format) ----------
{
  const rows = [
    ['Company'], [],
    ['Account Code','Account','Account Type','Debit - Year to date','Credit - Year to date'],
    ['21381','Annual Leave Provision','Liability', 0, 50000],
    ['11195','Workcover Prepaid','Asset', 4000, 0],
  ];
  const map = parseTrialBalance(rows);
  R.push(['MYOB: AL liability = 50000', approx(tbAccountBalance(map,'21381','liability'),50000)]);
  R.push(['MYOB: workcover asset = 4000', approx(tbAccountBalance(map,'11195','asset'),4000)]);
}

// ---------- Test 3: _tbDistribute ----------
departments = [{id:'site',name:'Site',xeroGroup:'Site Employees'},{id:'office',name:'Office',xeroGroup:'Office Employees'}];
{
  const prop = _tbDistribute(4000, {site:30000, office:10000}, 'prop');
  R.push(['distribute prop: site 3000', approx(prop.site,3000)]);
  R.push(['distribute prop: office 1000', approx(prop.office,1000)]);
  const one = _tbDistribute(4000, {site:30000, office:10000}, 'one', 'office');
  R.push(['distribute one: office 4000', approx(one.office,4000)]);
  R.push(['distribute one: site 0', approx(one.site,0)]);
  const zero = _tbDistribute(4000, {site:0, office:0}, 'prop');
  R.push(['distribute zero-weights: even 2000/2000', approx(zero.site,2000)&&approx(zero.office,2000)]);
}

// ---------- Test 4: AL allocation via applyTbAllocation (proportional) ----------
{
  setOpening('site',30000); setOpening('office',10000);   // sum 40000, TB 50000 -> +10000
  const variances = [{key:'annualLeaveProvision',name:'Annual Leave Provision',code:'21381',dept:'al',
                      tb:50000,inTb:true,expected:40000,computable:true,variance:10000,hasVariance:true}];
  openTbReconModal(variances);          // sets _tbVariancesPending, renders (must not throw)
  global.__radioChoice = { annualLeaveProvision: 'prop' };
  window._tbLslOpening = null;
  applyTbAllocation();
  R.push(['AL prop: site opening 37500', approx(getOpening('site'),37500)]);
  R.push(['AL prop: office opening 12500', approx(getOpening('office'),12500)]);
  R.push(['AL prop: openings sum to TB 50000', approx(getOpening('site')+getOpening('office'),50000)]);
  R.push(['AL prop: continueReconciliation called', global.__continued===true]);
}

// ---------- Test 5: AL allocation (charge to one department) ----------
{
  global.__continued = false;
  setOpening('site',30000); setOpening('office',10000);
  const variances = [{key:'annualLeaveProvision',name:'AL',code:'21381',dept:'al',
                      tb:50000,inTb:true,expected:40000,computable:true,variance:10000,hasVariance:true}];
  openTbReconModal(variances);
  global.__radioChoice = { annualLeaveProvision: 'one' };
  document.getElementById('tbAllocDept_annualLeaveProvision').value = 'office';
  applyTbAllocation();
  R.push(['AL one-dept: office opening 20000', approx(getOpening('office'),20000)]);
  R.push(['AL one-dept: site opening unchanged 30000', approx(getOpening('site'),30000)]);
}

// ---------- Test 6: LSL override (current has variance, non-current ties) ----------
{
  window._tbLslOpening = null;
  _tbLslResult = {
    employeeProvisions:[
      {department:'Site Employees',  provision:6000, isCurrent:true},
      {department:'Office Employees',provision:6000, isCurrent:true},
      {department:'Site Employees',  provision:4000, isCurrent:false},
    ],
    lastCurrentLiability:10000, lastNonCurrentLiability:4000
  };
  const variances = [
    {key:'lslCurrentProvision',name:'LSL Current',code:'21382',dept:'lslCurrent',
     tb:13000,inTb:true,expected:10000,computable:true,variance:3000,hasVariance:true},
    {key:'lslNonCurrentProvision',name:'LSL NonCur',code:'24800',dept:'lslNonCurrent',
     tb:4000,inTb:true,expected:4000,computable:true,variance:0,hasVariance:false},
  ];
  openTbReconModal(variances);
  global.__radioChoice = { lslCurrentProvision: 'prop' };
  applyTbAllocation();
  const ov = window._tbLslOpening;
  R.push(['LSL override exists', !!ov]);
  R.push(['LSL current: site 6500', ov && approx(ov.site.current,6500)]);
  R.push(['LSL current: office 6500', ov && approx(ov.office.current,6500)]);
  R.push(['LSL current sums to TB 13000', ov && approx(ov.site.current+ov.office.current,13000)]);
  R.push(['LSL non-current untouched: site 4000', ov && approx(ov.site.nonCurrent,4000)]);
  R.push(['LSL non-current untouched: office 0', ov && approx(ov.office.nonCurrent,0)]);
  R.push(['LSL non-current sums to TB 4000', ov && approx(ov.site.nonCurrent+ov.office.nonCurrent,4000)]);
}

// ---------- Test 7: processReconciliation gate — no TB → continues straight through ----------
{
  global.__continued = false;
  transactionData = [['x']];            // truthy so validation passes
  trialBalanceData = null;
  leaveLastSource = null;
  currentClientData = {};               // no lastReconciliation → no confirm
  window.lastTbRecon = 'stale';
  processReconciliation();
  R.push(['gate no-TB: continueReconciliation called', global.__continued===true]);
  R.push(['gate no-TB: lastTbRecon cleared to null', window.lastTbRecon===null]);
}

// ---------- Test 8: processReconciliation gate — TB present and ties → no modal, continues ----------
{
  global.__continued = false;
  transactionData = [['x']];
  document.getElementById('lsl_enabled').checked = false;   // skip LSL in the check
  setOpening('site',30000); setOpening('office',20000);      // AL opening sum 50000
  trialBalanceData = [
    ['Co'], ['Trial Balance'], [],
    ['Account','Debit','Credit'],
    ['21381 - Annual Leave Provision', null, 50000],          // ties exactly
  ];
  window._tbLslOpening = null;
  processReconciliation();
  R.push(['gate TB-ties: continueReconciliation called', global.__continued===true]);
  R.push(['gate TB-ties: tbRecon recorded (not applied)', !!window.lastTbRecon && window.lastTbRecon.applied===false]);
  const alRow = window.lastTbRecon && window.lastTbRecon.byAccount.find(a=>a.code==='21381');
  R.push(['gate TB-ties: AL row has no variance', !!alRow && alRow.hasVariance===false]);
}

global.__results = R;
`;

try { eval(scripts.join('\n;\n') + '\n;\n' + testSetup); }
catch (e) { console.log('EVAL ERROR:', e.message, '\n', e.stack); process.exit(2); }

const R = global.__results || [];
let fail = 0;
R.forEach(([label, ok]) => { console.log((ok?'PASS':'FAIL')+'  '+label); if(!ok) fail++; });
console.log(`\n${R.length-fail}/${R.length} passed`);
process.exit(fail ? 1 : 0);
