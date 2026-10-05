// 零依赖最小 XLSX 生成器（Node 内置 zlib，无需 npm 包）
const zlib = require('zlib');

const crc32 = buf => zlib.crc32(buf) >>> 0;

function zip(files) {
  const local = [], central = [];
  let offset = 0;
  files.forEach(f => {
    const nameBuf = Buffer.from(f.name, 'utf8');
    const comp = zlib.deflateRawSync(f.data, { level: 9 });
    const crc = crc32(f.data);

    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4);
    lh.writeUInt16LE(0x0800, 6); // UTF-8 文件名
    lh.writeUInt16LE(8, 8);      // deflate
    lh.writeUInt16LE(0, 10);
    lh.writeUInt16LE(0x21, 12);
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(comp.length, 18);
    lh.writeUInt32LE(f.data.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26);
    lh.writeUInt16LE(0, 28);
    local.push(lh, nameBuf, comp);

    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0x0800, 8);
    cd.writeUInt16LE(8, 10);
    cd.writeUInt16LE(0, 12);
    cd.writeUInt16LE(0x21, 14);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(comp.length, 20);
    cd.writeUInt32LE(f.data.length, 24);
    cd.writeUInt16LE(nameBuf.length, 28);
    cd.writeUInt16LE(0, 30);
    cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34);
    cd.writeUInt16LE(0, 36);
    cd.writeUInt32LE(0, 38);
    cd.writeUInt32LE(offset, 42);
    central.push(cd, nameBuf);

    offset += lh.length + nameBuf.length + comp.length;
  });

  const cdBuf = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cdBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([Buffer.concat(local), cdBuf, eocd]);
}

const xe = s => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  // eslint-disable-next-line no-control-regex
  .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

function colName(n) { // 1 -> A
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// sheet: { name, cols:[width,...], rows:[[cell,...],...] }
// cell: 字符串 / 数字 / {v, s}  s=1 表头 2=粗体 3=红字
function sheetXml(sheet) {
  const widths = (sheet.cols || []).map((w, i) =>
    `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');

  const rows = sheet.rows.map((row, ri) => {
    const cells = row.map((c, ci) => {
      if (c === null || c === undefined || c === '') return '';
      const ref = colName(ci + 1) + (ri + 1);
      let v = c, style = ri === 0 ? 1 : 0;
      if (typeof c === 'object') { v = c.v; style = c.s != null ? c.s : (ri === 0 ? 1 : 0); }
      if (typeof v === 'number' && isFinite(v)) {
        return `<c r="${ref}" s="${style}"><v>${v}</v></c>`;
      }
      return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xe(v)}</t></is></c>`;
    }).join('');
    return `<row r="${ri + 1}">${cells}</row>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="18"/>
${widths ? `<cols>${widths}</cols>` : ''}
<sheetData>${rows}</sheetData>
</worksheet>`;
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4">
<font><sz val="11"/><name val="宋体"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="宋体"/></font>
<font><b/><sz val="11"/><name val="宋体"/></font>
<font><sz val="11"/><color rgb="FFC0392B"/><name val="宋体"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFBFC9D6"/></left><right style="thin"><color rgb="FFBFC9D6"/></right><top style="thin"><color rgb="FFBFC9D6"/></top><bottom style="thin"><color rgb="FFBFC9D6"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="center" horizontal="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"><alignment vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

function safeSheetName(n, i) {
  let s = String(n || ('Sheet' + (i + 1))).replace(/[\\\/\?\*\[\]:]/g, '_');
  if (s.length > 31) s = s.slice(0, 31);
  return s || ('Sheet' + (i + 1));
}

/**
 * sheets: [{name, cols:[...], rows:[[...]]}]
 * 返回 Buffer（.xlsx）
 */
function build(sheets) {
  const files = [];
  const n = sheets.length;

  let overrides = '';
  let sheetTags = '';
  let rels = '';
  sheets.forEach((s, i) => {
    const nm = safeSheetName(s.name, i);
    overrides += `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    sheetTags += `<sheet name="${xe(nm)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
    rels += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: Buffer.from(sheetXml(s), 'utf8') });
  });
  rels += `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;

  files.push({
    name: '[Content_Types].xml',
    data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${overrides}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`, 'utf8')
  });
  files.push({
    name: '_rels/.rels',
    data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`, 'utf8')
  });
  files.push({
    name: 'xl/workbook.xml',
    data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheetTags}</sheets>
</workbook>`, 'utf8')
  });
  files.push({
    name: 'xl/_rels/workbook.xml.rels',
    data: Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`, 'utf8')
  });
  files.push({ name: 'xl/styles.xml', data: Buffer.from(STYLES, 'utf8') });

  return zip(files);
}

// CSV（带 BOM，Excel 直接打开不乱码）
function buildCsv(rows) {
  const q = v => {
    // 单元格可能是 {v, s} 样式对象，CSV 只取值
    const raw = (v && typeof v === 'object' && 'v' in v) ? v.v : v;
    const s = String(raw == null ? '' : raw);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return Buffer.from('\uFEFF' + rows.map(r => r.map(q).join(',')).join('\r\n'), 'utf8');
}

module.exports = { build, buildCsv };
