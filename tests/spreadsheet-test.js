/* Headless test harness: stubs browser APIs so we can load and exercise
   the Offline Spreadsheet script in Node. */
'use strict';
global.window = global;
global.addEventListener = ()=>{};
try { global.navigator.clipboard = {}; } catch(e) {
  Object.defineProperty(global, 'navigator', { value: { clipboard: {} }, configurable: true });
}
global.localStorage = (() => { const s = {}; return { getItem:k=>k in s?s[k]:null, setItem:(k,v)=>{s[k]=String(v);}, removeItem:k=>{delete s[k];} }; })();
global.confirm = () => true;
global.prompt = () => null;
global.alert = () => {};

let idCounter = 0;
class El {
  constructor(tag){ this.tagName=tag||'div'; this.id=''; this._class=new Set();
    this.style={}; this.dataset={}; this.children=[]; this.parentNode=null;
    this._listeners={}; this.value=''; this.textContent=''; this.innerText='';
    this.scrollTop=0; this.scrollLeft=0; this.clientWidth=800; this.clientHeight=600;
    this.attributes={};
    this.classList={ add:(c)=>this._class.add(c), remove:(c)=>this._class.delete(c),
      toggle:(c,f)=>{ if(f===undefined){ this._class.has(c)?this._class.delete(c):this._class.add(c);} else f?this._class.add(c):this._class.delete(c); },
      contains:(c)=>this._class.has(c) };
    this.focus=()=>{ global.document.activeElement=this; };
    this.blur=()=>{ if(global.document.activeElement===this)global.document.activeElement=null; };
    this.select=()=>{};
    this.setAttribute=(k,v)=>{this.attributes[k]=v; if(k==='id')this.id=v;};
    this.getAttribute=(k)=>this.attributes[k]??null;
    this.getBoundingClientRect=()=>({left:0,top:0,width:800,height:600,right:800,bottom:600});
    this.contains=(n)=>n===this;
    this.remove=()=>{ if(this.parentNode){const i=this.parentNode.children.indexOf(this); if(i>=0)this.parentNode.children.splice(i,1); this.parentNode=null;} };
  }
  addEventListener(){};
  set innerHTML(v){ this._inner=v; if(this._onSetHtml)this._onSetHtml(v); }
  get innerHTML(){ return this._inner||''; }
  appendChild(c){ this.children.push(c); c.parentNode=this; return c; }
  removeChild(c){ const i=this.children.indexOf(c); if(i>=0)this.children.splice(i,1); c.parentNode=null; return c; }
  insertBefore(c,ref){ const i=this.children.indexOf(ref); if(i<0)this.children.unshift(c); else this.children.splice(i,0,c); c.parentNode=this; return c; }
  querySelector(){ return null; }
  querySelectorAll(){ return []; }
}
global.document = {
  activeElement:null,
  createElement:(tag)=>new El(tag),
  createDocumentFragment:()=>{ const f=new El('fragment'); return f; },
  getElementById:(id)=>{ const e=new El('div'); e.id=id; return e; },
  querySelector:(sel)=>{ if(sel && sel[0]==='#') return global.document.getElementById(sel.slice(1)); return null; },
  querySelectorAll:()=>[],
  addEventListener:()=>{},
  body:new El('body'),
};
global.Blob=class Blob{constructor(p){this.p=p;}};
global.URL={createObjectURL:()=> 'blob:fake', revokeObjectURL:()=>{}};
global.FileReader=class{ readAsArrayBuffer(f){ this.result=f; if(this.onload)this.onload(); } };

// load SheetJS
global.XLSX = require(process.env.SHEETJS);
// load surgical OOXML engine
if (process.env.OOXMLJS) global.XLSXF = require(process.env.OOXMLJS);

// load the app script via vm so top-level vars land on the global object
const fs = require('fs');
const vm = require('vm');
vm.runInThisContext(fs.readFileSync(process.env.APPJS, 'utf8'), { filename: process.env.APPJS });

// ---- now test formulas ----
const t = (label, cond) => { console.log((cond?'PASS':'FAIL')+' - '+label); if(!cond) process.exitCode=1; };

function setCell(r,c,raw){ snap(); commitCell(r,c,raw); }

// basic grid values
newWorkbook();
setCell(0,0,'5');          // A1
setCell(0,1,'10');         // B1
setCell(0,2,'=A1*B1');     // C1 = 50
setCell(0,3,'=SUM(A1:B1)');// D1 = 15
setCell(0,4,'=IF(A1>3,"big","small")'); // E1 = big
setCell(0,5,'10%');        // F1 = 0.1 percent
setCell(1,0,'=A1+C9');     // A2 = 5 + empty(0) -> 5
setCell(1,1,'=1/0');       // B2 -> #DIV/0!

t('A1 == 5', resolve('Sheet1',0,0) === 5);
t('C1 == 50 (formula A1*B1)', resolve('Sheet1',0,2) === 50);
t('D1 == 15 (SUM range)', resolve('Sheet1',0,3) === 15);
t('E1 == "big" (IF)', resolve('Sheet1',0,4) === 'big');
t('F1 == 0.1 (percent parse)', resolve('Sheet1',0,5) === 0.1);
t('A2 formula on empty cell == 5', resolve('Sheet1',1,0) === 5);
t('B2 division error', isErr(resolve('Sheet1',1,1)) && resolve('Sheet1',1,1).__e === '#DIV/0!');

// cross-sheet reference
addSheet('Sheet2', true);
setCell(0,0,'42'); // A1 on Sheet2
setCell(0,1,'=Sheet1!C1'); // B1 on Sheet2 -> references Sheet1!C1 (50)
t('cross-sheet ref Sheet1!C1 == 50', resolve('Sheet2',0,1) === 50);

// VLOOKUP
active='Sheet1';
setCell(2,0,'Apple'); setCell(2,1,1);
setCell(3,0,'Banana');setCell(3,1,2);
setCell(2,2,'=VLOOKUP("Banana",A3:B4,2,FALSE)'); // 2
t('VLOOKUP Banana == 2', resolve('Sheet1',2,2) === 2);

// circular detection
setCell(4,0,'=B5'); setCell(4,1,'=A5'); // A5<->B5
t('circular ref -> #CYCLE!', (()=>{ const v=resolve('Sheet1',4,0); return isErr(v)&&v.__e==='#CYCLE!'; })());

// percent format check
const pc = getCell(0,5);
t('percent cell numfmt set', pc.numfmt === 'percent');

console.log('\n=== XLSX round-trip test ===');
// serialize current workbook state into an XLSX buffer, then import it back
const wb = XLSX.utils.book_new();
order.forEach((name)=>{
  const s = sheets[name];
  const ws = {};
  let maxR=0,maxC=0;
  Object.keys(s.cells).forEach((k)=>{const p=k.split(':'); maxR=Math.max(maxR,+p[0]); maxC=Math.max(maxC,+p[1]);});
  Object.keys(s.cells).forEach((k)=>{
    const p=k.split(':'); const r=+p[0], c=+p[1];
    const cell=s.cells[k]; const addr=XLSX.utils.encode_cell({r,c});
    const o={};
    if(cell.f){ o.f=cell.f; const v=resolve(name,r,c);
      if(!isErr(v)){ if(v instanceof Date){o.t='d';o.v=v;} else if(typeof v==='number'){o.t='n';o.v=v;} else if(typeof v==='boolean'){o.t='b';o.v=v;} else {o.t='s';o.v=v==null?'':String(v);} }
      else { o.t='s'; o.v='#'+v.__e; }
      if(cell.numfmt)o.z = {number:'#,##0.00',int:'0',percent:'0.00%',currency:'$#,##0.00',date:'yyyy-mm-dd'}[cell.numfmt];
    } else {
      if(typeof cell.v==='number'){o.t='n';o.v=cell.v;}
      else if(typeof cell.v==='boolean'){o.t='b';o.v=cell.v;}
      else if(cell.v instanceof Date){o.t='d';o.v=cell.v;}
      else {o.t='s';o.v=cell.v==null?'':String(cell.v);}
      if(cell.numfmt)o.z = {number:'#,##0.00',int:'0',percent:'0.00%',currency:'$#,##0.00',date:'yyyy-mm-dd'}[cell.numfmt];
    }
    ws[addr]=o;
  });
  if(maxR>=0){ ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:maxR||1,c:maxC||1}}); }
  XLSX.utils.book_append_sheet(wb, ws, name);
});
const out = XLSX.write(wb, {type:'buffer', bookType:'xlsx'});
console.log('xlsx bytes:', out.length);

// now read it back (simulating import)
const wb2 = XLSX.read(out, {type:'buffer', cellDates:true, cellNF:true});
refreshStateFromXLSX(wb2, 'roundtrip.xlsx');

t('roundtrip keeps 2 sheets', order.length === 2);
t('roundtrip A1 Sheet1 == 5', (()=>{ const c=sheets['Sheet1'].cells['0:0']; return c&&c.v===5; })());
t('roundtrip keeps C1 formula =A1*B1', (()=>{ const c=sheets['Sheet1'].cells['0:2']; return c&&c.f==='A1*B1'; })());
t('roundtrip keeps F1 percent cell', (()=>{ const c=sheets['Sheet1'].cells['0:5']; return c&&c.v===0.1 && c.numfmt==='percent'; })());
t('roundtrip keeps cross-sheet ref formula', (()=>{ const c=sheets['Sheet2'].cells['0:1']; return c&&c.f==='Sheet1!C1'; })());

// undo
active='Sheet1';
newWorkbook();
setCell(0,0,'hello');
snap();
setCell(0,1,'world');
const after = serialize();
undo();
t('undo reverts B1', (()=>{ const c=sheets[active].cells['0:1']; return !c; })());
console.log('DEBUG redostack entries:', redostack.length, JSON.stringify(redostack[0]||null).slice(0,80));
redo();
t('redo reapplies B1', (()=>{ const c=sheets[active].cells['0:1']; return c&&c.v==='world'; })());
console.log('DEBUG B1 after redo:', JSON.stringify(sheets[active].cells['0:1']||null));

// localStorage persist/restore
saveSessionNow();
const reread = loadSession();
t('serialize/deserialize roundtrip', reread && reread.sheets && Object.keys(reread.sheets).length>=1);

console.log('\nAll tests done.');