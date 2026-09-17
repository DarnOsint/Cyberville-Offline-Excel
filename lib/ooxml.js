/* ============================================================================
   ooxml.js — surgical OOXML / ZIP engine for Offline Spreadsheet
   ----------------------------------------------------------------------------
   Enables FULL-FIDELITY saving: the original .xlsx bytes are kept, and only
   the exact cells you edited are patched inside the real Excel XML. Anything
   you did NOT touch (charts, images, comments, styles, merged cells, ...)
   is carried through byte-for-byte.

   Components:
     - CRC32
     - Raw DEFLATE decompressor (RFC 1951), used to read zip parts we rewrite
     - ZIP parser + builder (verbatim copy of untouched entries, STORE for
       rebuilt entries)
     - OOXML helpers: parse sheet XML into head/tail/row/cell markup, patch
       workbook.xml / rels / [Content_Types].xml / styles.xml
   Pure JS, no DOM, works in browser and Node.
   ============================================================================ */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.XLSXF = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ------------------------------------------------------------------ UTF-8 */
  var TE = (typeof TextEncoder !== 'undefined') ? new TextEncoder() : null;
  var TD = (typeof TextDecoder !== 'undefined') ? new TextDecoder('utf-8') : null;
  function str2u8(s) {
    if (TE) return TE.encode(s);
    var b = new Uint8Array(s.length * 2), n = 0;
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c < 128) b[n++] = c;
      else if (c < 2048) { b[n++] = 0xC0 | (c >> 6); b[n++] = 0x80 | (c & 63); }
      else { b[n++] = 0xE0 | (c >> 12); b[n++] = 0x80 | ((c >> 6) & 63); b[n++] = 0x80 | (c & 63); }
    }
    return b.slice(0, n);
  }
  function u8str(u8) {
    if (TD) return TD.decode(u8);
    var s = '', i = 0;
    while (i < u8.length) {
      var a = u8[i++];
      if (a < 128) s += String.fromCharCode(a);
      else if (a < 224 && i < u8.length) s += String.fromCharCode(((a & 31) << 6) | (u8[i++] & 63));
      else if (i < u8.length + 1) s += String.fromCharCode(((a & 15) << 12) | ((u8[i++] & 63) << 6) | (u8[i++] & 63));
    }
    return s;
  }

  /* ------------------------------------------------------------------ CRC32 */
  var CRC_TABLE = null;
  function crcTable() {
    if (CRC_TABLE) return CRC_TABLE;
    CRC_TABLE = new Int32Array(256);
    for (var i = 0; i < 256; i++) {
      var c = i;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      CRC_TABLE[i] = c;
    }
    return CRC_TABLE;
  }
  function crc32(u8) {
    var t = crcTable(), c = 0xFFFFFFFF;
    for (var i = 0; i < u8.length; i++) c = t[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /* ------------------------------------------------------------- bit reader */
  function BitReader(bytes) { this.b = bytes; this.i = 0; this.cache = 0; this.n = 0; }
  BitReader.prototype.read = function (n) {
    var out = 0, shift = 0;
    while (n > 0) {
      if (this.n <= 0) { if (this.i >= this.b.length) throw new Error('inflate: unexpected end'); this.cache = this.b[this.i++]; this.n = 8; }
      var take = Math.min(this.n, n);
      out |= (this.cache & ((1 << take) - 1)) << shift;
      this.cache >>>= take;
      this.n -= take; n -= take; shift += take;
    }
    return out;
  };

  /* ---------------------------------------------------------------- inflate */
  var LEN_BASE = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
  var LEN_EXTRA = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
  var DIST_BASE = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
  var DIST_EXTRA = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
  var CODE_ORDER = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];

  function fixedLens(n, specs) {
    var a = new Array(n);
    for (var i = 0; i < specs.length; i += 3) {
      var from = specs[i], to = specs[i + 1], len = specs[i + 2];
      for (var x = from; x <= to; x++) a[x] = len;
    }
    return a;
  }
  var FIXED_LIT = fixedLens(288, [0,143,8, 144,255,9, 256,279,7, 280,287,8]);
  var FIXED_DIST = fixedLens(32, [0,31,5]);

  /* Canonical Huffman table (zlib style): counts per length + symbols in
     code order. decodeSym reads exactly the code length it needs. */
  function buildHuffman(lengths) {
    var i, counts = new Array(16).fill(0);
    for (i = 0; i < lengths.length; i++) counts[lengths[i]]++;
    if (counts[0] === lengths.length) return null;
    var left = 1;
    for (i = 1; i <= 15; i++) { left = (left << 1) - counts[i]; if (left < 0) return null; }
    var symbols = [];
    for (var len = 1; len <= 15; len++) {
      for (i = 0; i < lengths.length; i++) if (lengths[i] === len) symbols.push(i);
    }
    var first = new Array(16), total = 0;
    for (len = 1; len <= 15; len++) { first[len] = total; total += counts[len]; }
    return { counts: counts, first: first, symbols: symbols };
  }
  function decodeSym(br, h) {
    if (!h) throw new Error('inflate: empty huffman table');
    var code = 0, index = 0, first = 0;
    for (var len = 1; len <= 15; len++) {
      code |= br.read(1);
      var count = h.counts[len];
      if (code - first < count) return h.symbols[index + (code - first)];
      index += count;
      first = (first + count) << 1;
      code <<= 1;
    }
    return -1;
  }

  function inflate(raw) {
    var br = new BitReader(raw), out = [], last = 0;
    while (!last) {
      last = br.read(1);
      var type = br.read(2);
      if (type === 0) {
        br.n = 0;
        var len = br.read(16), nlen = br.read(16);
        for (var i = 0; i < len; i++) out.push(br.read(8));
      } else if (type === 1 || type === 2) {
        var lit, dist;
        if (type === 1) { lit = buildHuffman(FIXED_LIT); dist = buildHuffman(FIXED_DIST); }
        else {
          var hlit = br.read(5) + 257, hdist = br.read(5) + 1, hclen = br.read(4) + 4;
          var cl = new Int32Array(19), j;
          for (j = 0; j < hclen; j++) cl[CODE_ORDER[j]] = br.read(3);
          var ct = buildHuffman(cl);
          var lens = [], total = hlit + hdist;
          while (lens.length < total) {
            var s = decodeSym(br, ct);
            if (s < 16) lens.push(s);
            else if (s === 16) { var rep = br.read(2) + 3, prev = lens.length ? lens[lens.length - 1] : 0; while (rep--) lens.push(prev); }
            else if (s === 17) { rep = br.read(3) + 3; while (rep--) lens.push(0); }
            else { rep = br.read(7) + 11; while (rep--) lens.push(0); }
          }
          lit = buildHuffman(lens.slice(0, hlit));
          dist = buildHuffman(lens.slice(hlit));
        }
        for (;;) {
          var sym = decodeSym(br, lit);
          if (sym < 256) out.push(sym);
          else if (sym === 256) break;
          else {
            var li = sym - 257, length = LEN_BASE[li] + br.read(LEN_EXTRA[li]);
            var ds = decodeSym(br, dist);
            if (ds > 29) throw new Error('inflate: bad distance code');
            var distv = DIST_BASE[ds] + br.read(DIST_EXTRA[ds]);
            if (distv > out.length) throw new Error('inflate: distance too far back');
            var start = out.length - distv;
            for (var z = 0; z < length; z++) out.push(out[start + z]);
          }
        }
      } else throw new Error('inflate: invalid block type');
    }
    return Uint8Array.from(out);
  }

  /* -------------------------------------------------------------------- zip */
  function readU16(b, o) { return b[o] | (b[o + 1] << 8); }
  function readU32(b, o) { return ((b[o] | (b[o + 1] << 8) | (b[o + 2] << 16)) + (b[o + 3] * 0x1000000)) >>> 0; }
  function writeU16(b, o, v) { b[o] = v & 255; b[o + 1] = (v >> 8) & 255; }
  function writeU32(b, o, v) { b[o] = v & 255; b[o + 1] = (v >> 8) & 255; b[o + 2] = (v >> 16) & 255; b[o + 3] = (v >>> 24) & 255; }

  function findEOCD(bytes) {
    var max = bytes.length - 22;
    if (max < 0) throw new Error('not a zip (too small)');
    var start = Math.max(0, max - 65536);
    for (var i = max; i >= start; i--) {
      if (bytes[i] === 0x50 && bytes[i + 1] === 0x4B && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) return i;
    }
    throw new Error('not a zip (no end-of-central-directory)');
  }

  function zipParse(bytes) {
    var eocd = findEOCD(bytes);
    var cdOff = readU32(bytes, eocd + 16), count = readU16(bytes, eocd + 10);
    var entries = [], pos = cdOff;
    for (var i = 0; i < count; i++) {
      if (readU32(bytes, pos) !== 0x02014B50) break;
      var method = readU16(bytes, pos + 10);
      var crc = readU32(bytes, pos + 16);
      var compSize = readU32(bytes, pos + 20), uncompSize = readU32(bytes, pos + 24);
      var nameLen = readU16(bytes, pos + 28), extraLen = readU16(bytes, pos + 30), commentLen = readU16(bytes, pos + 32);
      var localOff = readU32(bytes, pos + 42);
      var name = u8str(bytes.subarray(pos + 46, pos + 46 + nameLen));
      var lNameLen = readU16(bytes, localOff + 26), lExtraLen = readU16(bytes, localOff + 28);
      entries.push({
        name: name, method: method, crc: crc,
        compSize: compSize, uncompSize: uncompSize,
        dataStart: localOff + 30 + lNameLen + lExtraLen
      });
      pos += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  function findEntry(entries, name) {
    for (var i = 0; i < entries.length; i++) if (entries[i].name === name) return entries[i];
    return null;
  }

  function entryData(bytes, e) {
    var slice = bytes.subarray(e.dataStart, e.dataStart + e.compSize);
    return e.method === 0 ? slice : inflate(slice);
  }

  function entryText(bytes, e) { return u8str(entryData(bytes, e)); }

  function toBytes(d) {
    if (typeof d === 'string') return str2u8(d);
    if (d instanceof Uint8Array) return d;
    if (Array.isArray(d)) return Uint8Array.from(d);
    return d;
  }

  function zipBuild(outEntries) {
    var parts = [], central = [], localOff = 0, cdLen = 0;
    outEntries.forEach(function (e) {
      var nameBytes = str2u8(e.name);
      var data = toBytes(e.data);
      var method = e._raw ? e.method : 0;
      var compSize = e._raw ? e.compSize : data.length;
      var uncompSize = e._raw ? e.uncompSize : data.length;
      var crc = e._raw ? e.crc : crc32(data);
      if (e._raw && crc == null) throw new Error('raw entry needs crc: ' + e.name);
      var lh = new Uint8Array(30 + nameBytes.length);
      writeU32(lh, 0, 0x04034B50); writeU16(lh, 4, 20); writeU16(lh, 6, 0x0800);
      writeU16(lh, 8, method); writeU16(lh, 10, 0); writeU16(lh, 12, 0);
      writeU32(lh, 14, crc); writeU32(lh, 18, compSize); writeU32(lh, 22, uncompSize);
      writeU16(lh, 26, nameBytes.length); writeU16(lh, 28, 0);
      lh.set(nameBytes, 30);
      parts.push(lh); parts.push(data);
      var ch = new Uint8Array(46 + nameBytes.length);
      writeU32(ch, 0, 0x02014B50); writeU16(ch, 4, 20); writeU16(ch, 6, 20);
      writeU16(ch, 8, 0x0800); writeU16(ch, 10, method);
      writeU16(ch, 12, 0); writeU16(ch, 14, 0);
      writeU32(ch, 16, crc); writeU32(ch, 20, compSize); writeU32(ch, 24, uncompSize);
      writeU16(ch, 28, nameBytes.length); writeU16(ch, 30, 0); writeU16(ch, 32, 0);
      writeU16(ch, 34, 0); writeU16(ch, 36, 0); writeU32(ch, 38, 0); writeU32(ch, 42, localOff);
      ch.set(nameBytes, 46);
      central.push(ch); cdLen += ch.length;
      localOff += lh.length + data.length;
    });
    var eo = new Uint8Array(22);
    writeU32(eo, 0, 0x06054B50); writeU16(eo, 4, 0); writeU16(eo, 6, 0);
    writeU16(eo, 8, outEntries.length); writeU16(eo, 10, outEntries.length);
    writeU32(eo, 12, cdLen); writeU32(eo, 16, localOff); writeU16(eo, 20, 0);
    var total = localOff + cdLen + 22, buf = new Uint8Array(total), off = 0;
    parts.forEach(function (p) { buf.set(p, off); off += p.length; });
    central.forEach(function (p) { buf.set(p, off); off += p.length; });
    buf.set(eo, off);
    return buf;
  }

  /* ----------------------------------------------------------- XML helpers */
  function escAttr(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function escText(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function getAttr(tag, name) {
    var m = new RegExp('\\b' + name + '="([^"]*)"').exec(tag);
    return m ? m[1] : null;
  }
  function countOf(xml, tag) {
    var m = new RegExp('<' + tag + '\\b[^>]*?count="(\\d+)"').exec(xml);
    return m ? +m[1] : null;
  }
  function maxNumFmtId(xml) {
    var max = 0, re = /numFmt numFmtId="(\d+)"/g, m;
    while ((m = re.exec(xml))) max = Math.max(max, +m[1]);
    return max;
  }

  /* Appends newInner into a <tag count="N"> ... </tag> container.
     Handles: open/close pair, self-closing tag, or missing container.
     count = number of NEW records being injected (bumped onto existing count). */
  function containerAppendMany(xml, tag, newInner, count, insertBefore) {
    var openRe = new RegExp('<' + tag + '\\b([^>]*?)(\\/?)>');
    var m = openRe.exec(xml);
    if (!m) {
      var idx = -1;
      if (insertBefore) for (var c = 0; c < insertBefore.length; c++) { var ii = xml.indexOf(insertBefore[c]); if (ii >= 0) { idx = ii; break; } }
      var block = '<' + tag + ' count="' + count + '">' + newInner + '</' + tag + '>';
      return idx < 0 ? xml + block : xml.slice(0, idx) + block + xml.slice(idx);
    }
    var attrs = m[1], selfClose = m[2] === '/';
    var newCount = (countOf(xml, tag) || 0) + count;
    attrs = attrs.replace(/count="\d*"/, 'count="' + newCount + '"');
    var newOpen = '<' + tag + attrs + '>';
    var after = m.index + m[0].length;
    if (selfClose) return xml.slice(0, m.index) + newOpen + newInner + '</' + tag + '>' + xml.slice(after);
    var closeIdx = xml.indexOf('</' + tag + '>', after);
    if (closeIdx < 0) return xml;
    return xml.slice(0, m.index) + newOpen + xml.slice(after, closeIdx) + newInner + xml.slice(closeIdx);
  }

  /* ---------------------------------------------------------------- sheet */
  /* Splits a sheet XML into: head (everything through '<sheetData' openers),
     tail (after '</sheetData>'), and the raw <sheetData> body string. */
  function sheetParts(xml) {
    var si = xml.indexOf('<sheetData');
    if (si < 0) return { head: xml, body: '', tail: '', isEmpty: true };
    var openEnd = xml.indexOf('>', si);
    if (openEnd < 0) return { head: xml, body: '', tail: '', isEmpty: true };
    var openTag = xml.slice(si, openEnd + 1);
    var selfClose = /\/>$/.test(openTag);
    var close = xml.indexOf('</sheetData>', openEnd + 1);
    var body = '', tail = '';
    if (selfClose) { tail = xml.slice(openEnd + 1); }
    else if (close >= 0) { body = xml.slice(openEnd + 1, close); tail = xml.slice(close + 12); }
    else { body = xml.slice(openEnd + 1); }
    return { head: xml.slice(0, si), body: body, tail: tail, isEmpty: body.length === 0 };
  }

  function findTagEnd(s, start) {
    var q = false;
    for (var i = start; i < s.length; i++) {
      var c = s[i];
      if (c === '"') q = !q;
      else if (c === '>' && !q) return i + 1;
    }
    return s.length;
  }

  /* Extracts full <row>...</row> and <c>...</c> markup from a sheetData body. */
  function parseRows(body) {
    var rows = {}, i = 0;
    while (i < body.length) {
      var ri = body.indexOf('<row', i);
      if (ri < 0) break;
      var gt = findTagEnd(body, ri);
      if (gt <= ri) break;
      var open = body.slice(ri, gt);
      var r = (parseInt(getAttr(open, 'r') || '0', 10)) || 1;
      r = r - 1; /* normalize to 0-based grid rows to match app convention */
      if (open[gt - 2] === '/') { if (!rows[r]) rows[r] = { tag: open, cells: {} }; i = gt; continue; }
      var nri = body.indexOf('<row', gt);
      var contentEnd = nri < 0 ? body.length : nri;
      var cells = parseCells(body.slice(gt, contentEnd));
      if (!rows[r]) rows[r] = { tag: open, cells: cells };
      else for (var k in cells) rows[r].cells[k] = cells[k];
      i = contentEnd;
    }
    return rows;
  }

  function parseCells(content) {
    var cells = {}, i = 0;
    while (i < content.length) {
      var ci = content.indexOf('<c', i);
      if (ci < 0) break;
      var gt = findTagEnd(content, ci);
      if (gt <= ci) break;
      var open = content.slice(ci, gt);
      var addr = getAttr(open, 'r');
      var full;
      if (open[gt - 2] === '/') { full = open; i = gt; }
      else {
        var end = content.indexOf('</c>', gt);
        if (end < 0) { full = open; i = gt; }
        else { full = open + content.slice(gt, end) + '</c>'; i = end + 4; }
      }
      if (addr) cells[addr] = full;
    }
    return cells;
  }

  function addrRC(addr) {
    var m = /^([A-Z]+)(\d+)$/.exec(addr);
    if (!m) return null;
    var col = 0;
    for (var i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
    return { r: +m[2] - 1, c: col - 1 };
  }

  /* ------------------------------------------------------------------ wb  */
  /* Returns [{name, rid, sheetId}] in file order. */
  function parseWorkbookSheets(wbXml) {
    var out = [], re = /<sheet\b([^>]*)\/?>/g, m;
    while ((m = re.exec(wbXml))) {
      var attrs = m[1];
      out.push({ name: getAttr(attrs, 'name'), rid: getAttr(attrs, 'r:id'), sheetId: +getAttr(attrs, 'sheetId') });
    }
    return out;
  }

  /* Returns map rid -> target. */
  function parseRels(relsXml) {
    var map = {}, re = /<Relationship\b([^>]*)\/?>/g, m;
    while ((m = re.exec(relsXml))) {
      var attrs = m[1], id = getAttr(attrs, 'Id'), target = getAttr(attrs, 'Target');
      if (id) map[id] = target || '';
    }
    return map;
  }

  function resolveTarget(target) {
    if (!target) return null;
    if (target.charAt(0) === '/') return target.replace(/^\//, '');
    return 'xl/' + target.replace(/^\.\.\//, '');
  }

  function rebuildSheetsBlock(wbXml, sheets) {
    var inner = '';
    sheets.forEach(function (s, idx) {
      inner += '<sheet name="' + escAttr(s.name) + '" sheetId="' + (s.sheetId || (idx + 1)) + '" r:id="' + s.rid + '"/>';
    });
    var block = '<sheets>' + inner + '</sheets>';
    var m = /<sheets\b[\s\S]*?<\/sheets>/.exec(wbXml);
    if (m) return wbXml.slice(0, m.index) + block + wbXml.slice(m.index + m[0].length);
    var si = wbXml.indexOf('<sheets');
    if (si >= 0) {
      var ee = wbXml.indexOf('>', si);
      var e2 = wbXml.indexOf('>', ee + 1);
      return wbXml.slice(0, si) + block + wbXml.slice((e2 >= 0 ? e2 : ee) + 1);
    }
    var close = wbXml.lastIndexOf('</workbook>');
    if (close >= 0) return wbXml.slice(0, close) + block + wbXml.slice(close);
    return wbXml + block;
  }

  function maxRid(relsXml) {
    var max = 0, re = /Id="rId(\d+)"/g, m;
    while ((m = re.exec(relsXml))) max = Math.max(max, +m[1]);
    return max;
  }

  /* Rewrites the rels so that only Relationships pointing at kept sheet parts
     remain; appends relationships for brand-new sheet parts. */
  function rewriteRels(relsXml, usedParts, allocated) {
    var kept = [], re = /<Relationship\b[^>]*\/?>/g, m;
    var usedResolved = {};
    for (var k in usedParts) usedResolved[k] = 1;
    re.lastIndex = 0;
    while ((m = re.exec(relsXml))) {
      var attrs = m[1], target = getAttr(attrs, 'Target');
      var resolved = resolveTarget(target);
      if (resolved && usedResolved[resolved]) { kept.push(m[0]); }
      else {
        /* a rel whose sheet part was deleted is dropped ONLY if it is a worksheet rel */
        if (/worksheet/.test(getAttr(attrs, 'Type') || '')) { /* drop */ }
        else kept.push(m[0]);
      }
    }
    var next = maxRid(relsXml || '');
    allocated.forEach(function (a) {
      var pn = (typeof a === 'string') ? a : a.partName;
      var rid = a.rid ? a.rid : 'rId' + (++next);
      kept.push('<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' + pn + '"/>');
    });
    var out = relsXml;
    var openM = /<Relationships\b[^>]*>/.exec(out);
    var head = out.slice(0, openM.index + openM[0].length), tail = out.slice(openM.index + openM[0].length);
    var inner = '';
    kept.forEach(function (rel) { inner += rel; });
    return head + inner + '</Relationships>' + (tail.indexOf('</Relationships>') >= 0 ? tail.slice(tail.indexOf('</Relationships>') + 15) : tail);
  }

  /* [Content_Types].xml: remove overrides for deleted sheet parts, add for new. */
  function rewriteContentTypes(ctXml, removedParts, allocated) {
    var m = /<Types\b[^>]*>/;
    var head, inner = '';
    var om = m.exec(ctXml);
    head = ctXml.slice(0, om.index + om[0].length);
    var childrenRe = /<(Default|Override)\b[^>]*\/?>/g, c;
    while ((c = childrenRe.exec(ctXml))) {
      if (c.index > om.index + om[0].length) {
        var tag = c[0];
        var part = getAttr(tag, 'PartName');
        if (part && removedParts[part.replace(/^\//, '')]) continue;
        inner += tag;
      }
    }
    allocated.forEach(function (a) {
      var pn = (typeof a === 'string') ? a : a.partName;
      inner += '<Override PartName="/xl/worksheets/' + pn + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    });
    return head + inner + '</Types>';
  }

  /* styles.xml patching -------------------------------------------------- */
  function freshStylesXml(STY) {
    var font = '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
    var fill = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>';
    var border = '<border><left/><right/><top/><bottom/><diagonal/></border>';
    var cellStyleXf = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>';
    var pr = collectStyleFamilies(STY);
    var num = pr.nums.length ? '<numFmts count="' + pr.nums.length + '">' + pr.nums.join('') + '</numFmts>' : '';
    var fonts = '<fonts count="' + (1 + pr.fonts.length) + '">' + font + pr.fonts.join('') + '</fonts>';
    var fills = '<fills count="' + (2 + pr.fills.length) + '">' + fill + pr.fills.join('') + '</fills>';
    var borders = '<borders count="1">' + border + '</borders>';
    var css = '<cellStyleXfs count="1">' + cellStyleXf + '</cellStyleXfs>';
    var cxfs = '<cellXfs count="' + (1 + pr.xfs.length) + '">' + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' + pr.xfs.join('') + '</cellXfs>';
    var cst = '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac">'
      + num + fonts + fills + borders + css + cxfs + cst
      + '<dxfs count="0"/><tableStyles count="0"/></styleSheet>';
  }

  function styleXfXml(st) {
    var x = '<xf numFmtId="' + st.numFmtId + '" fontId="' + st.fontId + '" fillId="' + st.fillId + '" borderId="0" xfId="0"';
    if (st.numFmtId !== 0) x += ' applyNumberFormat="1"';
    if (st.fontId !== 0) x += ' applyFont="1"';
    if (st.fillId !== 0) x += ' applyFill="1"';
    if (st.align) x += ' applyAlignment="1"><alignment horizontal="' + st.align + '"/></xf>';
    else x += '/>';
    return x;
  }

  function collectStyleFamilies(STY) {
    var nums = [], fonts = [], fills = [], xfs = [];
    for (var code in STY.nums) nums[STY.nums[code]] = '<numFmt numFmtId="' + STY.nums[code] + '" formatCode="' + escAttr(code) + '"/>';
    for (var fk in STY.fonts) fonts[STY.fonts[fk]] = collectFontXml(fk);
    for (var fi in STY.fills) fills[STY.fills[fi]] = '<fill><patternFill patternType="solid"><fgColor rgb="' + fi + '"/><bgColor indexed="64"/></patternFill></fill>';
    STY.order.forEach(function (st) { xfs.push(styleXfXml(st)); });
    var sort = function (a) { var o = []; for (var i = 0; i < a.length; i++) if (a[i]) o.push(a[i]); return o; };
    return { nums: sort(nums), fonts: sort(fonts), fills: sort(fills), xfs: sort(xfs) };
  }

  function collectFontXml(fontKey) {
    var p = fontKey.indexOf(':');
    var flags = p < 0 ? fontKey : fontKey.slice(0, p);
    var color = p < 0 ? '' : fontKey.slice(p + 1);
    var f = '<font>';
    if (flags.indexOf('B') >= 0) f += '<b/>';
    if (flags.indexOf('I') >= 0) f += '<i/>';
    if (flags.indexOf('U') >= 0) f += '<u/>';
    f += '<sz val="11"/>';
    if (color) {
      /* accept '#rrggbb' or 'rrggbb' -> 'FFrrggbb' */
      var hex = color.replace(/^#/, '');
      f += '<color rgb="FF' + hex.toUpperCase() + '"/>';
    }
    f += '<name val="Calibri"/></font>';
    return f;
  }

  /* Patches an EXISTING styles.xml to add registered styles. Returns new xml. */
  function patchStyles(existing, STY) {
    if (!STY.order.length) return existing;
    var xml = existing;
    var pr = collectStyleFamilies(STY);
    var before = ['<cellXfs', '<cellStyleXfs', '<fonts', '<numFmts'];
    if (pr.nums.length) {
      /* skip any numFmt whose formatCode already exists (prevents duplicates) */
      var keep = [];
      pr.nums.forEach(function (entry) {
        var m = /formatCode="([^"]*)"/.exec(entry);
        if (m && existing.indexOf('formatCode="' + m[1] + '"') >= 0) return;
        keep.push(entry);
      });
      if (keep.length) xml = containerAppendMany(xml, 'numFmts', keep.join(''), keep.length, before);
    }
    if (pr.fonts.length) xml = containerAppendMany(xml, 'fonts', pr.fonts.join(''), pr.fonts.length, ['<cellStyleXfs', '<cellXfs', '<fills', '<border']);
    if (pr.fills.length) xml = containerAppendMany(xml, 'fills', pr.fills.join(''), pr.fills.length, ['<cellStyleXfs', '<cellXfs', '<borders']);
    if (pr.xfs.length) xml = containerAppendMany(xml, 'cellXfs', pr.xfs.join(''), pr.xfs.length, ['<cellStyles']);
    return xml;
  }

  /* Counts families already present (for index allocation when patching). */
  function stylesBase(existing) {
    var b = { nums: 0, fonts: 1, fills: 2, xfs: 1, nextNumId: 164 };
    if (!existing) return b;
    b.nums = countOf(existing, 'numFmts') || 0;
    b.fonts = countOf(existing, 'fonts') || 1;
    b.fills = countOf(existing, 'fills') || 2;
    b.xfs = countOf(existing, 'cellXfs') || 1;
    b.nextNumId = Math.max(164, (maxNumFmtId(existing) || 0) + 1);
    return b;
  }

  return {
    crc32: crc32,
    inflate: inflate,
    zipParse: zipParse,
    zipBuild: zipBuild,
    findEntry: findEntry,
    entryData: entryData,
    entryText: entryText,
    sheetParts: sheetParts,
    parseRows: parseRows,
    parseCells: parseCells,
    addrRC: addrRC,
    escAttr: escAttr,
    escText: escText,
    getAttr: getAttr,
    parseWorkbookSheets: parseWorkbookSheets,
    parseRels: parseRels,
    resolveTarget: resolveTarget,
    rebuildSheetsBlock: rebuildSheetsBlock,
    rewriteRels: rewriteRels,
    rewriteContentTypes: rewriteContentTypes,
    stylesBase: stylesBase,
    patchStyles: patchStyles,
    freshStylesXml: freshStylesXml,
    collectFontXml: collectFontXml,
    styleXfXml: styleXfXml,
    containerAppendMany: containerAppendMany,
    _stylesBase: stylesBase,
    NS_R: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    NS_PKG_REL: 'http://schemas.openxmlformats.org/package/2006/relationships',
    DEFAULT_CT_OVERRIDES: '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
  };
});