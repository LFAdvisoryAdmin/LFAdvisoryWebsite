// Harness: run payroll-history.html's export pipeline against Zummo's real snapshot
const fs = require('fs');

// ---- minimal DOM stub ----
const elements = {};
function makeEl(id) {
  return {
    id, value: '', innerHTML: '', style: {}, textContent: '',
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    appendChild(){}, querySelector(){ return null; }, querySelectorAll(){ return []; },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
  };
}
const outputDiv = makeEl('output');
// journal table parsed from innerHTML on demand
outputDiv.querySelector = (sel) => {
  if (!outputDiv.innerHTML.includes('<table')) return null;
  // crude tbody row parser
  const rowsHtml = [...outputDiv.innerHTML.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)];
  const bodyRows = rowsHtml.filter(r => !r[1].includes('<th'));
  return {
    querySelectorAll: (s) => bodyRows.map(r => {
      const cells = [...r[1].matchAll(/<td[^>]*>([\s\S]*?)<\/td>/g)].map(c => ({
        innerText: c[1].replace(/<[^>]*>/g,''), colSpan: /colspan/.test(r[1]) ? 2 : 1
      }));
      return {
        classList: { contains: (cls) => r[0].includes('highlight') },
        querySelectorAll: () => cells
      };
    })
  };
};
global.document = {
  getElementById(id) { if (id === 'output') return outputDiv; if (!elements[id]) elements[id] = makeEl(id); return elements[id]; },
  createElement(tag) { return makeEl(tag); },
  querySelector(sel) { return sel === '#output table' ? outputDiv.querySelector(sel) : null; },
  querySelectorAll() { return []; },
  body: { appendChild(){} },
  addEventListener(){},
};
global.window = global;
global.addEventListener = () => {};
global.removeEventListener = () => {};
global.setInterval = () => 0;
global.alert = (msg) => console.log('ALERT:', String(msg).split('\n')[0]);
global.localStorage = { getItem(){return null;}, setItem(){}, removeItem(){} };
global.location = { hash: '', origin: '', pathname: '' };
global.sessionStorage = global.localStorage;
global.fetch = () => Promise.resolve({ ok:false, json:()=>({}) });
global.XLSX = null;

// ---- recording jsPDF mock ----
const texts = [];
let pageNum = 1;
class FakePDF {
  constructor(){ }
  setFontSize(){} setFont(){} setLineWidth(){} setDrawColor(){} setTextColor(){} setFillColor(){}
  line(){} rect(){} roundedRect(){}
  addPage(){ pageNum++; }
  getTextWidth(t){ return String(t).length * 1.8; }
  text(txt, x, y){ (Array.isArray(txt)?txt:[txt]).forEach(t => texts.push({ page: pageNum, txt: String(t) })); }
  splitTextToSize(t){ return [t]; }
  save(name){ console.log('PDF SAVED:', name, '| pages:', pageNum); }
}
global.jspdf = { jsPDF: FakePDF };
global.window.jspdf = global.jspdf;

// ---- load app code ----
const code = fs.readFileSync(__dirname + '/history-scripts.js', 'utf8');
try { eval(code); } catch (e) { console.log('TOP-LEVEL EVAL ERROR:', e.message); }

// ---- feed Zummo April 2026 snapshot ----
const h = JSON.parse(fs.readFileSync('C:/Users/f869f/OneDrive - LF Advisory Pty Ltd/LF Advisory Workpapers/lf-advisory-payroll-history.json','utf8'));
const snapshot = h['Zummo Juicers Pty Ltd']['April 2026'];

try {
  exportAuditPDF(snapshot);
} catch (e) {
  console.log('EXPORT THREW:', e.message, '\n', e.stack.split('\n').slice(0,4).join('\n'));
}

// ---- report which sections made it into the PDF ----
console.log('\n--- Page 5 (LSL provision) full text ---');
texts.filter(t => t.page === 5).forEach(t => console.log(`p5: ${t.txt}`));
console.log('\n--- Page 9 (Appendix D) full text ---');
texts.filter(t => t.page === 9).forEach(t => console.log(`p9: ${t.txt}`));
console.log('\nTotal text calls:', texts.length);
