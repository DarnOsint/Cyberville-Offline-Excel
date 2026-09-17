const zlib = require('zlib');
const XLSXF = require('/Users/macbook/ExcelOffline/lib/ooxml.js');

let pass = 0, fail = 0;
function t(name, cond) { if (cond) { pass++; } else { fail++; console.log('FAIL:', name); } }

// 1) inflate vs zlib on random byte arrays (raw deflate)
const sizes = [0, 1, 10, 100, 1000, 10000, 100000, 500000];
for (const sz of sizes) {
  const data = Buffer.alloc(sz);
  for (let i = 0; i < sz; i++) data[i] = Math.floor(Math.random() * 256);
  const def = zlib.deflateRawSync(data);
  const out = XLSXF.inflate(def);
  t('inflate random size ' + sz, Buffer.compare(Buffer.from(out), data) === 0);
}

// 2) inflate on textual data (small deflate, matches sheets better)
for (let passNo = 0; passNo < 5; passNo++) {
  let txt = '';
  for (let i = 0; i < 20000; i++) txt += (i % 5 === 0) ? '<sheetData><row r="' + i + '"/>' : 'hello world & ' + i + ';';
  const buf = Buffer.from(txt, 'utf8');
  const def = zlib.deflateRawSync(buf, { level: 6 });
  const out = XLSXF.inflate(def);
  t('inflate text ' + passNo, Buffer.from(out).toString('utf8') === txt);
}

// 3) zip build -> zip parse roundtrip with mixed store + deflate entries
function zipRoundtrip() {
  const entries = [
    { name: '[Content_Types].xml', data: Buffer.from('<Types/>'), method: 0 },
    { name: 'docProps/app.xml', data: Buffer.from('hello app'), method: 0 },
  ];
  // deflate entry: fake the raw compressed bytes
  const rawTxt = Buffer.from('x'.repeat(5000));
  const deflated = zlib.deflateRawSync(rawTxt);
  entries.push({ name: 'xl/worksheets/sheet1.xml', data: deflated, method: 8, compSize: deflated.length, uncompSize: rawTxt.length, crc: XLSXF.crc32(rawTxt), _raw: true });
  const built = XLSXF.zipBuild(entries.map(e => Object.assign({}, e)));
  const parsed = XLSXF.zipParse(built);
  t('zip parse count', parsed.length === 3);
  const xlsf = parsed.find(e => e.name === 'xl/worksheets/sheet1.xml');
  const decoded = XLSXF.entryData(built, xlsf);
  t('deflate entry inflates', Buffer.compare(Buffer.from(decoded), rawTxt) === 0);
  const ct = parsed.find(e => e.name === '[Content_Types].xml');
  t('store entry reads', XLSXF.entryText(built, ct) === '<Types/>');
}
zipRoundtrip();

// 4) verbatim copy fidelity: rebuild must reproduce identical bytes for a store entry
{
  const txt = Buffer.from('some bytes to roundtrip exactly...');
  const { zipBuild, zipParse, entryData } = XLSXF;
  const built = zipBuild([{ name: 'x.txt', data: txt, method: 0 }]);
  const parsed = zipParse(built);
  const again = zipBuild([{ name: 'x.txt', data: entryData(built, parsed[0]), method: 0 }]);
  t('verbatim store rebuild identical', Buffer.compare(again, built) === 0);
}

// 5) sheetParts / parseRows / parseCells
{
  const xml = '<?xml version="1.0"?><worksheet><dimension ref="A1:B2"/><sheetData><row r="1" ht="20"><c r="A1" s="3"><v>10</v></c><c r="B1" t="s"><v>hello</v></c></row><row r="2"><c r="A2" s="5" t="inlineStr"><is><t xml:space="preserve">x &amp; y</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>';
  const parts = XLSXF.sheetParts(xml);
  t('sheetParts head', parts.head.indexOf('<sheetData') === -1 && parts.head.indexOf('<dimension') >= 0);
  t('sheetParts tail keeps mergeCells', parts.tail.indexOf('mergeCells') >= 0);
  const rows = XLSXF.parseRows(parts.body);
  t('parseRows r1 cell count', Object.keys(rows[0].cells).length === 2);
  t('parseRows r1 (0-based)', !!rows[1] && rows[1].cells['A2'].indexOf('inlineStr') >= 0);
  t('parseRows r0 tag preserved ht', /ht="20"/.test(rows[0].tag));
}

// 6) styles patch
{
  const base = '<?xml version="1.0"?><styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts><fonts count="1"><font></font></fonts><fills count="2"><fill></fill></fills><cellXfs count="1"><xf/></cellXfs></styleSheet>';
  const STY = {
    order: [{ numFmtId: 165, fontId: 1, fillId: 2, align: 'center' }],
    nums: { '#,##0.00': 165 },
    fonts: { 'B:': 1 },
    fills: { 'FFFF0000': 2 }
  };
  const out = XLSXF.patchStyles(base, STY);
  t('patchStyles numFmts count bumped', /<numFmts count="2">/.test(out) && /#,##0.00/.test(out));
  t('patchStyles cellXfs count bumped', /<cellXfs count="2">/.test(out));
  t('patchStyles xf added', /numFmtId="165" fontId="1" fillId="2"/.test(out));
}

// 7) fresh styles + workbook helpers
{
  const STY = { order: [], nums: {}, fonts: {}, fills: {} };
  const f = XLSXF.freshStylesXml(STY);
  t('fresh styles valid-ish', /<cellXfs count="1">/.test(f) && /<\/styleSheet>/.test(f));
  const wb = XLSXF.rebuildSheetsBlock('<workbook><sheets><sheet name="Old" sheetId="1" r:id="rId1"/></sheets></workbook>', [{ name: 'New', rid: 'rId1', sheetId: 1 }]);
  t('rebuildSheetsBlock renamed', wb.indexOf('name="New"') >= 0 && wb.indexOf('Old') === -1);
}

console.log(`\nPASS: ${pass}  FAIL: ${fail}`);
process.exit(fail ? 1 : 0);
