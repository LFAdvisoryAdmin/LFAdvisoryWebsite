// Harness: run payroll-reconciliation.html's restore + audit export against Zummo's real snapshot
const fs = require('fs');

// ---- minimal DOM stub ----
const elements = {};
function makeEl(id) {
  const el = {
    id, value: '', innerHTML: '', style: {}, textContent: '', checked: false,
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    appendChild(){}, remove(){},
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    scrollIntoView(){},
    querySelector(){ return null; }, querySelectorAll(){ return []; },
  };
  return el;
}
const outputDiv = makeEl('output');
function parseTable() {
  if (!outputDiv.innerHTML.includes('<table')) return null;
  const rowsHtml = [...outputDiv.innerHTML.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)];
  const bodyRows = rowsHtml.filter(r => !r[1].includes('<th'));
  return {
    querySelectorAll: () => bodyRows.map(r => {
      const cells = [...r[1].matchAll(/<td[^>]*>([\s\S]*?)<\/td>/g)].map(c => ({
        innerText: c[1].replace(/<[^>]*>/g,'').trim(), colSpan: /colspan/.test(r[0]+c[0]) ? 2 : 1
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
global.window = global;
global.addEventListener = () => {};
global.removeEventListener = () => {};
global.setInterval = () => 0;
global.setTimeout = (fn) => 0;   // don't run deferred UI callbacks
global.alert = (msg) => console.log('ALERT:', String(msg).split('\n')[0]);
global.confirm = () => false;
global.localStorage = { getItem(){return null;}, setItem(){}, removeItem(){} };
global.sessionStorage = global.localStorage;
global.location = { hash: '', origin: '', pathname: '', search: '' };
global.fetch = () => Promise.resolve({ ok:false, json:()=>({}) });
global.navigator = { userAgent: 'test' };
global.XLSX = null;

// ---- recording jsPDF mock ----
const texts = [];
let pageNum = 1;
class FakePDF {
  setFontSize(){} setFont(){} setLineWidth(){} setDrawColor(){} setTextColor(){} setFillColor(){}
  line(){} rect(){} roundedRect(){}
  addPage(){ pageNum++; }
  getTextWidth(t){ return String(t).length * 1.8; }
  text(txt, x, y){ (Array.isArray(txt)?txt:[txt]).forEach(t => texts.push({ page: pageNum, txt: String(t) })); }
  splitTextToSize(t){ return [t]; }
  save(name){ console.log('PDF SAVED:', name, '| pages:', pageNum); }
}
global.jspdf = { jsPDF: FakePDF };

// ---- load recon app code ----
const html = fs.readFileSync('C:/Users/f869f/LFAdvisoryWebsite/payroll-reconciliation.html','utf8');
const scripts = [...html.matchAll(/<script(?![^>]*src)[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]);
console.log('script blocks:', scripts.length);
try { eval(scripts.join('\n;\n')); } catch (e) { console.log('TOP-LEVEL EVAL ERROR:', e.message); }

// ---- restore Zummo April 2026 snapshot the way loadPastPeriod does ----
const h = JSON.parse(fs.readFileSync('C:/Users/f869f/OneDrive - LF Advisory Pty Ltd/LF Advisory Workpapers/lf-advisory-payroll-history.json','utf8'));
const snapshot = h['Zummo Juicers Pty Ltd']['April 2026'];

try {
  restoreActiveRec(snapshot);
  console.log('after restore: window.lastLSLData set?', !!global.lastLSLData || !!global.window.lastLSLData);
  console.log('journal table present?', !!parseTable());
} catch (e) {
  console.log('RESTORE THREW:', e.message);
}


// ---- XLSX stub + Excel export ----
global.XLSX = {
  utils: {
    book_new: () => ({ SheetNames: [], Sheets: {} }),
    aoa_to_sheet: (rows) => ({ __rows: rows }),
    book_append_sheet: (wb, ws, name) => { wb.SheetNames.push(name); wb.Sheets[name] = ws; },
    sheet_add_aoa: (ws, rows) => { ws.__rows = (ws.__rows || []).concat(rows); },
    encode_cell: () => 'A1',
    decode_range: () => ({ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }),
  },
  writeFile: (wb, name) => {
    console.log('XLSX SAVED:', name);
    console.log('sheets:', wb.SheetNames.join(' | '));
    wb.SheetNames.filter(n => /lsl|appendix|summary/i.test(n)).forEach(n => {
      console.log('');
      console.log('=== SHEET: ' + n + ' ===');
      (wb.Sheets[n].__rows || []).slice(0, 45).forEach(r => {
        const line = (r || []).map(c => (c && typeof c === 'object') ? (c.v !== undefined ? c.v : (c.f ? '=' + c.f : '')) : c).join(' | ');
        if (line.replace(/[\s|]/g, '')) console.log(line);
      });
    });
  }
};
try { generateExcelAuditReport(); } catch (e) { console.log('EXCEL EXPORT THREW:', e.message); console.log(e.stack.split('\n').slice(0,3).join('\n')); }
