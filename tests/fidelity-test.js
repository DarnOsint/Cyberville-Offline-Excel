/* E2E fidelity test: import a RICH xlsx (values, formulas, merges, drawings,
   charts, comments, styles), edit a few cells, save surgically, and verify:
   - untouched parts are byte-identical
   - edited cells are patched
   - merged cells survive
   - the result is a valid zip + well-formed XML and re-reads in SheetJS
   Also exercises the fresh-workbook (no import) save path. */
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
    this.attributes={}; this.click=()=>{};
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
  addEventListener:()=>{},
  body:new El('body'),
};
global.Blob=class Blob{constructor(p){this.p=p;}};
global.URL={createObjectURL:()=> 'blob:fake', revokeObjectURL:()=>{}};
global.FileReader=class{ readAsArrayBuffer(f){ this.result=f; if(this.onload)this.onload(); } };

const fs = require('fs');
const vm = require('vm');

global.XLSX = require(process.env.SHEETJS);
global.XLSXF = require(process.env.OOXMLJS);

vm.runInThisContext(fs.readFileSync(process.env.APPJS, 'utf8'), { filename: process.env.APPJS });

const t = (label, cond) => { console.log((cond?'PASS':'FAIL')+' - '+label); if(!cond) process.exitCode=1; };

/* ---------------- build a rich source xlsx as raw OOXML ---------------- */
const TE = (s)=>new TextEncoder().encode(s);
function store(name, xml){ return { name, data: TE(typeof xml==='string'?xml:xml.data), method: 0, _rawfalse:true }; }

const sheet1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:D5"/>
<sheetViews><sheetView workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols><col min="1" max="1" width="12" customWidth="1"/></cols>
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="n" s="1"><v>42</v></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>Hello</t></is></c><c r="C2" s="2"><f>B1*2</f><v>84</v></c></row>
<row r="3"><c r="A3" s="3"><v>99</v></c></row>
<row r="5"><c r="D5" s="2"><v>7.5</v></c></row>
</sheetData>
<mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="D2:E4"/></mergeCells>
<drawing r:id="rId1"/>
<legacyDrawing r:id="rId2"/>
</worksheet>`;

const sheet2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:B1"/>
<sheetViews><sheetView workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<sheetData>
<row r="1"><c r="A1" t="n"><v>1</v></c><c r="B1" t="n"><v>2</v></c></row>
</sheetData>
</worksheet>`;

const drawing = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<xdr:oneCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:ext cx="100000" cy="100000"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="PIC"/><xdr:cNvPicPr/></xdr:nvPicPr>
<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>
<xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;

const chart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart><c:plotArea><c:layout/><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/></c:numCache></c:numRef></c:val></c:ser></c:lineChart></c:plotArea></c:chart><c:printSettings/></c:chartSpace>`;

const comments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>DarnOsint</author></authors><commentList><comment ref="B1" authorId="0"><text><t>secret note</t></text></comment></commentList></comments>`;

const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"><font><b/></font></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"><font><i/></font></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFill="1"><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets>
<calcPr calcId="191029"/>
</workbook>`;

const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

const sheet1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>`;

const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>`;

const sharedStrings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>A1val</t></si></sst>`;

const png = Uint8Array.from([137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,0,0,1,0,0,0,1,0,8,6,0,0,0,31,21,196,137,0,0,0,4,103,65,77,65,0,0,177,143,11,252,97,5,0,0,0,9,112,72,89,115,0,0,14,195,0,0,14,195,1,199,111,168,100,0,0,0,13,73,68,65,84,120,1,99,248,207,192,0,0,0,6,0,42,0,0,0,0,0,0]);
const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

function mkStore(name,data){ return { name, data, method:0 }; }
const srcEntries = [
  mkStore('[Content_Types].xml', contentTypes),
  mkStore('_rels/.rels', rootRels),
  mkStore('xl/workbook.xml', workbook),
  mkStore('xl/_rels/workbook.xml.rels', wbRels),
  mkStore('xl/styles.xml', styles),
  mkStore('xl/sharedStrings.xml', sharedStrings),
  mkStore('xl/worksheets/sheet1.xml', sheet1),
  mkStore('xl/worksheets/sheet2.xml', sheet2),
  mkStore('xl/_rels/worksheets/sheet1.xml.rels', sheet1Rels),
  mkStore('xl/drawings/drawing1.xml', drawing),
  mkStore('xl/drawings/_rels/drawing1.xml.rels', drawingRels),
  mkStore('xl/charts/chart1.xml', chart),
  mkStore('xl/comments1.xml', comments),
  mkStore('xl/media/image1.png', png),
];
const src = XLSXF.zipBuild(srcEntries);
console.log('source xlsx bytes:', src.length);
fs.writeFileSync('/tmp/fidelity-src.xlsx', Buffer.from(src));

/* record byte content of parts we must NOT touch */
function entryBytes(z, name){
  const ents = XLSXF.zipParse(z);
  const found = ents.filter(f => f.name === name);
  let tgt = ents.find(u => u.name.indexOf('drawing')>=0);
  if(found.length===0) console.log('DEBUG strict miss; sample:', JSON.stringify(tgt&&tgt.name), 'tgt chars:', (tgt&&tgt.name.split('').map(ch=>ch.charCodeAt(0))).slice(0,14), '| asked:', JSON.stringify(name), 'asked chars:', name.split('').map(ch=>ch.charCodeAt(0)));
  const e = found.length? found[0] : null;
  return Array.from(z.subarray(e.dataStart, e.dataStart+e.compSize)); }
const origDrawing = entryBytes(src,'xl/drawings/drawing1.xml');
const origChart  = entryBytes(src,'xl/charts/chart1.xml');
const origPng    = entryBytes(src,'xl/media/image1.png');
const origComments = entryBytes(src,'xl/comments1.xml');

/* ---------------- simulate import (as index.html does) ---------------- */
const wb = XLSX.read(src, {type:'array', cellDates:true, cellNF:true});
refreshStateFromXLSX(wb, 'rich.xlsx', src);

t('OO captured on import', OO !== null);
t('both sheets have oo context', !!sheets['Sheet1'].oo && !!sheets['Sheet2'].oo);
t('merges captured (2)', sheets['Sheet1'].merges.length === 2);
t('imported B1 value 42', (()=>{const c=sheets['Sheet1'].cells['0:1']; return c && c.v===42;})());
t('imported C2 formula B1*2', (()=>{const c=sheets['Sheet1'].cells['1:2']; return c && c.f==='B1*2';})());

/* ---- edit: change A1 (was shared string), clear B1 (was styled),
        set new formula C4 = B1+8 -> 50, new text E5, bold a fresh cell */
active='Sheet1';
// A1 = 42 (number replaces shared string)
commitCell(0,0,'42');
t('A1 dirty after edit', !!sheets['Sheet1'].dirty['0:0']);
// clear B1
commitCell(0,1,'');
t('B1 cleared & cell removed', !sheets['Sheet1'].cells['0:1'] && !!sheets['Sheet1'].cleared['0:1']);
// new formula (A1 edited to 42 -> =A1+8 resolves to 50)
commitCell(3,2,'=A1+8');
t('C4 formula resolves 50', resolve('Sheet1',3,2)===50);
commitCell(4,4,'plain text');
commitCell(5,0,'7');
applyFmt('b',1);

/* ---- surgical save ---- */
const out = buildSurgicalXLSX();
console.log('saved xlsx bytes:', out.length);
fs.writeFileSync('/tmp/fidelity-out.xlsx', Buffer.from(out));
t('bold style registered during save', STY!==null && STY.order.length>=1);

/* verify parts byte-identical */
function entryBytes2(arr, name){ const e=XLSXF.zipParse(arr); const f=XLSXF.findEntry(e,name); return Array.from(arr.subarray(f.dataStart,f.dataStart+f.compSize)); }
t('drawing1.xml byte-identical', JSON.stringify(entryBytes2(out,'xl/drawings/drawing1.xml'))===JSON.stringify(origDrawing));
t('chart1.xml byte-identical', JSON.stringify(entryBytes2(out,'xl/charts/chart1.xml'))===JSON.stringify(origChart));
t('image1.png byte-identical', JSON.stringify(entryBytes2(out,'xl/media/image1.png'))===JSON.stringify(origPng));
t('comments1.xml byte-identical', JSON.stringify(entryBytes2(out,'xl/comments1.xml'))===JSON.stringify(origComments));

/* inspect sheet1 xml */
const outEntries = XLSXF.zipParse(out);
const s1e = XLSXF.findEntry(outEntries,'xl/worksheets/sheet1.xml');
const s1xml = XLSXF.entryText(out, s1e);
t('A1 became numeric 42', /<c r="A1"[^>]*><v>42<\/v><\/c>/.test(s1xml));
t('B1 removed', !/<c r="B1"/.test(s1xml));
t('C4 authored formula =A1+8 with v 50', /<c r="C4"[^>]*><f>A1\+8<\/f><v>50<\/v><\/c>/.test(s1xml));
t('E5 authored inlineStr text', /<c r="E5"[^>]*t="inlineStr"><is><t xml:space="preserve">plain text<\/t><\/is><\/c>/.test(s1xml));
t('C2 original formula preserved', /<c r="C2" s="2"><f>B1\*2<\/f><v>84<\/v><\/c>/.test(s1xml));
t('merge cells preserved', /<mergeCells count="2"><mergeCell ref="A1:C1"\/><mergeCell ref="D2:E4"\/><\/mergeCells>/.test(s1xml));
t('drawing rel kept', /<drawing r:id="rId1"\/>/.test(s1xml));
t('dimension updated', /<dimension ref="A1:E6"\/>/.test(s1xml));
const a3xml = /<c r="A3" s="3"><v>99<\/v><\/c>/.test(s1xml);
t('A3 untouched styled cell preserved', a3xml);
const d5xml = /<c r="D5" s="2"><v>7.5<\/v><\/c>/.test(s1xml);
t('D5 untouched preserved', d5xml);

/* styles.xml patched for new bold xf */
const stE = XLSXF.findEntry(outEntries,'xl/styles.xml');
const stylesXml = XLSXF.entryText(out, stE);
const boldNew = /<cellXfs[^>]*count="5"/.test(stylesXml);
t('cellXfs grew to 5 (bold injected)', boldNew);

/* workbook.xml still has both sheets */
const wbE2 = XLSXF.findEntry(outEntries,'xl/workbook.xml');
const wbXml2 = XLSXF.entryText(out, wbE2);
t('workbook has 2 sheets', /<sheet .*name="Sheet1".*\/><sheet .*name="Sheet2".*\/>/.test(wbXml2));

/* ---- read back with SheetJS ---- */
const back = XLSX.read(out, {type:'array', cellDates:true, cellNF:true});
t('read-back Sheet1 A1 == 42', back.Sheets['Sheet1'].A1 && back.Sheets['Sheet1'].A1.v===42);
t('read-back keeps Sheet2 A1 == 1', back.Sheets['Sheet2'].A1 && back.Sheets['Sheet2'].A1.v===1);
t('read-back merges preserved (2)', back.Sheets['Sheet1']['!merges'] && back.Sheets['Sheet1']['!merges'].length===2);

/* ---- fresh (no import) save path ---- */
newWorkbook();
commitCell(0,0,'alpha');
commitCell(0,1,'=A1');
applyFmt('center','center');
const out2 = buildFreshXLSX();
fs.writeFileSync('/tmp/fidelity-fresh.xlsx', Buffer.from(out2));
const back2 = XLSX.read(out2, {type:'array', cellNF:true});
t('fresh workbook read-back A1 alpha', back2.Sheets['Sheet1'].A1 && back2.Sheets['Sheet1'].A1.v==='alpha');
t('fresh workbook formula kept', back2.Sheets['Sheet1'].B1 && back2.Sheets['Sheet1'].B1.f==='A1');

console.log('\nAll tests done.');