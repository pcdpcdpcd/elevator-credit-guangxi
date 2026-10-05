/**
 * 单位信息导入解析（零依赖）
 * 支持两种文件：
 *   .csv / .txt —— 逗号分隔文本（自动识别 ; 和制表符，自动去 BOM）
 *   .xlsx       —— Excel 工作簿（自行解开 zip 并读取第一张工作表）
 * 只返回「表格二维数组」，字段名映射与落库交给调用方。
 */
const zlib = require('zlib');

// ---------------- CSV ----------------
/** 解析分隔符文本为二维数组 */
function parseDelimited(text) {
  let s = String(text == null ? '' : text).replace(/^\uFEFF/, '');
  // 探测分隔符：取第一行，看哪种符号出现得多
  const firstLine = s.split(/\r?\n/)[0] || '';
  const cand = [',', ';', '\t'];
  let sep = ',';
  let best = -1;
  cand.forEach(c => {
    const n = (firstLine.match(new RegExp(c === '\t' ? '\\t' : '\\' + c, 'g')) || []).length;
    if (n > best) { best = n; sep = c; }
  });

  const rows = [];
  let row = [], cur = '', inQ = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (inQ) {
      if (ch === '"') {
        if (s[i + 1] === '"') { cur += '"'; i++; } else inQ = false;
      } else cur += ch;
    } else {
      if (ch === '"') inQ = true;
      else if (ch === sep) { row.push(cur); cur = ''; }
      else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
      else if (ch === '\r') { /* 忽略 */ }
      else cur += ch;
    }
  }
  row.push(cur); rows.push(row);
  // 去掉全空行，并去掉尾部空列
  return rows
    .map(r => r.map(x => String(x == null ? '' : x).trim()))
    .filter(r => r.some(x => x !== ''));
}

// ---------------- XLSX ----------------
function xmlDecode(s) {
  return String(s)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}
function colToIdx(letters) {
  let n = 0;
  for (const ch of String(letters).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
/** 从 zip 里取出指定文件（xlsx 本质是一个 zip 包） */
function zipRead(buf, want) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0 && i > buf.length - 70000; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('文件不是有效的 xlsx（读不到 zip 目录）');
  const total = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  let fallback = null;
  for (let i = 0; i < total; i++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) break;
    const method = buf.readUInt16LE(p + 10);
    const compSize = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const offset = buf.readUInt32LE(p + 42);
    const fname = buf.slice(p + 46, p + 46 + nameLen).toString('utf8');
    const isSheet = /^xl\/worksheets\/sheet\d+\.xml$/.test(fname);
    if (fname === want || (want === 'sheet1' && isSheet)) {
      const lNameLen = buf.readUInt16LE(offset + 26);
      const lExtraLen = buf.readUInt16LE(offset + 28);
      const start = offset + 30 + lNameLen + lExtraLen;
      const data = buf.slice(start, start + compSize);
      const out = method === 0 ? data : zlib.inflateRawSync(data);
      if (fname === want) return out;
      if (!fallback) fallback = { name: fname, data: out };
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  return fallback ? fallback.data : null;
}
function sharedStrings(buf) {
  const xml = zipRead(buf, 'xl/sharedStrings.xml');
  if (!xml) return [];
  const s = xml.toString('utf8');
  const out = [];
  const re = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = re.exec(s))) {
    const ts = m[1].match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || [];
    out.push(ts.map(x => xmlDecode(x.replace(/^<t\b[^>]*>/, '').replace(/<\/t>$/, ''))).join(''));
  }
  return out;
}
/** 解析 xlsx 第一张工作表为二维数组 */
function parseXlsx(buf) {
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
  const shared = sharedStrings(buf);
  const xml = zipRead(buf, 'xl/worksheets/sheet1.xml') || zipRead(buf, 'sheet1');
  if (!xml) throw new Error('xlsx 里没有找到工作表');
  const s = xml.toString('utf8');
  const rows = [];
  const rowRe = /<row\b[^>]*>([\s\S]*?)<\/row>/g;
  let rm;
  while ((rm = rowRe.exec(s))) {
    const cells = [];
    const cRe = /<c\b([^>]*?)\/?>([\s\S]*?)<\/c>|<c\b([^>]*?)\/>/g;
    let cm;
    while ((cm = cRe.exec(rm[1]))) {
      const attrs = cm[1] || cm[3] || '';
      const inner = cm[2] || '';
      const ref = /r="([A-Za-z]+)(\d+)"/.exec(attrs);
      const t = /t="([^"]+)"/.exec(attrs);
      let v = '';
      const vm = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner);
      if (t && t[1] === 's') {
        v = vm ? (shared[parseInt(vm[1], 10)] || '') : '';
      } else if (t && t[1] === 'inlineStr') {
        const tm = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(inner);
        v = tm ? xmlDecode(tm[1]) : '';
      } else {
        v = vm ? xmlDecode(vm[1]) : '';
      }
      const col = ref ? colToIdx(ref[1]) : cells.length;
      cells[col] = String(v == null ? '' : v).trim();
    }
    rows.push(cells);
  }
  return rows
    .map(r => r.map(x => String(x == null ? '' : x).trim()))
    .filter(r => r.some(x => x !== ''));
}

// ---------------- 表头 → 字段 ----------------
/** 表头别名：左侧为系统字段，右侧为允许的表头写法 */
const FIELD_ALIASES = {
  name: ['单位名称', '维保单位名称', '维保单位', '企业名称', '公司名称', '单位', 'name'],
  jurisdiction: ['辖区', '办公点辖区', '所属辖区', '区县', '所在城区', 'jurisdiction'],
  license_no: ['许可证号', '许可证编号', '资质证书号', '证书编号', 'license_no'],
  level: ['许可级别', '级别', '资质等级', '等级', 'level'],
  scope: ['许可范围', '业务范围', '经营范围', 'scope'],
  office: ['办公地点', '办公地址', '地址', '住所', '注册地址', 'office'],
  contact: ['联系人', '负责人', '法人', 'contact'],
  phone: ['联系电话', '电话', '手机号', '联系方式', 'phone'],
  note: ['备注', '说明', 'note']
};
const REQUIRED_FIELD = 'name';

function normHeader(h) {
  return String(h == null ? '' : h).replace(/\s/g, '').replace(/[（(].*?[)）]/g, '').toLowerCase();
}
/** 把表头行翻译成字段数组；返回 {map:[field|null], missing:bool} */
function mapHeader(headerRow) {
  const map = headerRow.map(h => {
    const n = normHeader(h);
    if (!n) return null;
    for (const field of Object.keys(FIELD_ALIASES)) {
      if (FIELD_ALIASES[field].some(a => normHeader(a) === n)) return field;
    }
    // 包含式匹配：例如「维保单位名称（必填）」
    for (const field of Object.keys(FIELD_ALIASES)) {
      if (FIELD_ALIASES[field].some(a => n.indexOf(normHeader(a)) >= 0)) return field;
    }
    return null;
  });
  return { map, missing: !map.includes(REQUIRED_FIELD) };
}

module.exports = { parseDelimited, parseXlsx, mapHeader, FIELD_ALIASES };
