/* xlsx_lite.js —— 无依赖的 .xlsx 写出（store 压缩 + inlineStr + 公式 + 简单样式）
   浏览器: window.XLSXLite ; node: module.exports
   用法: XLSXLite.build([{name:'参数', cols:[6,20,12], rows:[[{v:'序号',s:1}, ...]], freeze:1}], {definedNames:{侧板厚:'参数!$C$6'}}) -> Uint8Array
   单元格: null | '文本' | 数字 | {v, f(公式), s(样式0-7), fmt}                     */
(function (root) {
  const CRC = (() => { const t = new Int32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c; } return t; })();
  function crc32(u8) { let c = -1; for (let i = 0; i < u8.length; i++) c = (c >>> 8) ^ CRC[(c ^ u8[i]) & 0xFF]; return (c ^ -1) >>> 0; }
  const enc = new TextEncoder();
  const U = s => enc.encode(s);

  function zip(files) {
    const parts = [], cds = []; let off = 0;
    const now = new Date();
    const dTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xFFFF;
    const dDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;
    for (const f of files) {
      const nb = U(f.name), data = f.data, size = data.length, crc = crc32(data);
      const lh = new Uint8Array(30 + nb.length), dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true); dv.setUint16(4, 20, true); dv.setUint16(6, 0x0800, true);
      dv.setUint16(8, 0, true); dv.setUint16(10, dTime, true); dv.setUint16(12, dDate, true);
      dv.setUint32(14, crc, true); dv.setUint32(18, size, true); dv.setUint32(22, size, true);
      dv.setUint16(26, nb.length, true); dv.setUint16(28, 0, true);
      lh.set(nb, 30);
      parts.push(lh, data);
      const ch = new Uint8Array(46 + nb.length), cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true); cv.setUint16(10, 0, true);
      cv.setUint16(12, dTime, true); cv.setUint16(14, dDate, true);
      cv.setUint32(16, crc, true); cv.setUint32(20, size, true); cv.setUint32(24, size, true);
      cv.setUint16(28, nb.length, true); cv.setUint16(30, 0, true); cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true); cv.setUint16(36, 0, true); cv.setUint32(38, 0, true);
      cv.setUint32(42, off, true); ch.set(nb, 46);
      cds.push(ch);
      off += lh.length + size;
    }
    const cdSize = cds.reduce((s, c) => s + c.length, 0);
    const eocd = new Uint8Array(22), ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, files.length, true); ev.setUint16(10, files.length, true);
    ev.setUint32(12, cdSize, true); ev.setUint32(16, off, true);
    const out = new Uint8Array(off + cdSize + 22); let p = 0;
    for (const b of parts) { out.set(b, p); p += b.length; }
    for (const c of cds) { out.set(c, p); p += c.length; }
    out.set(eocd, p);
    return out;
  }

  const X = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const COL = i => { let s = '', n = i + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const num = v => (Math.abs(v) >= 1e-12 || v === 0) ? String(+v.toFixed(10)) : '0';

  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="0.0000"/><numFmt numFmtId="165" formatCode="0.00"/></numFmts>
<fonts count="4">
<font><sz val="11"/><name val="\u7b49\u7ebf"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="\u7b49\u7ebf"/></font>
<font><b/><sz val="12"/><name val="\u7b49\u7ebf"/></font>
<font><b/><sz val="11"/><name val="\u7b49\u7ebf"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2"><border/>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="8">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  function cellXml(ref, c) {
    if (c == null || c === '') return '';
    if (typeof c === 'object' && c.f) return `<c r="${ref}"${c.s ? ` s="${c.s}"` : ''}><f>${X(c.f)}</f></c>`;
    const v = (typeof c === 'object') ? c.v : c;
    const s = (typeof c === 'object' && c.s) ? ` s="${c.s}"` : '';
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number' && isFinite(v)) return `<c r="${ref}"${s}><v>${num(v)}</v></c>`;
    return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${X(v)}</t></is></c>`;
  }

  function sheetXml(sh) {
    let s = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`;
    if (sh.freeze) s += `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${sh.freeze}" topLeftCell="A${sh.freeze + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`;
    if (sh.cols && sh.cols.length) s += `<cols>` + sh.cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + `</cols>`;
    s += `<sheetData>`;
    (sh.rows || []).forEach((row, ri) => {
      const r = ri + 1;
      const cells = (row || []).map((c, ci) => cellXml(COL(ci) + r, c)).join('');
      s += `<row r="${r}">${cells}</row>`;
    });
    s += `</sheetData></worksheet>`;
    return s;
  }

  function build(sheets, opts) {
    opts = opts || {};
    const files = [];
    const n = sheets.length;
    let ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`;
    sheets.forEach((sh, i) => { ct += `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`; });
    ct += `</Types>`;
    files.push({ name: '[Content_Types].xml', data: U(ct) });
    files.push({ name: '_rels/.rels', data: U(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`) });

    let wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    sheets.forEach((sh, i) => { wbRels += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`; });
    wbRels += `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: U(wbRels) });

    let wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>`;
    sheets.forEach((sh, i) => { wb += `<sheet name="${X(sh.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`; });
    wb += `</sheets>`;
    const dn = opts.definedNames || {};
    const keys = Object.keys(dn);
    if (keys.length) {
      wb += `<definedNames>` + keys.map(k => `<definedName name="${X(k)}">${X(dn[k])}</definedName>`).join('') + `</definedNames>`;
    }
    wb += `<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>`;
    files.push({ name: 'xl/workbook.xml', data: U(wb) });
    files.push({ name: 'xl/styles.xml', data: U(STYLES) });
    sheets.forEach((sh, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: U(sheetXml(sh)) }));
    return zip(files);
  }

  const api = { build };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.XLSXLite = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
