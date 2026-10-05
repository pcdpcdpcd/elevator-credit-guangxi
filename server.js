const http = require('http');
const fs = require('fs');
const path = require('path');
const XLSX = require('./xlsx');
const IMP = require('./import');

// ---------- 读取用户配置（config.js，可选）----------
let userCfg = {};
try { userCfg = require('./config') || {}; } catch (e) { userCfg = {}; }

// 端口：环境变量 > config.js > 默认 3000
const PORT = process.env.PORT || userCfg.port || 3000;
const DB_FILE = path.join(__dirname, 'db.json');
const HTML_FILE = path.join(__dirname, 'index.html');

// 加分项目封顶：管理创新5 / 智慧管理10 / 保险服务10 / 公益服务10
const BONUS_CAP = { '加分1': 5, '加分2': 10, '加分3': 10, '加分4': 10 };
const QUARTER_LABEL = { 1: '第一季度', 2: '第二季度', 3: '第三季度', 4: '第四季度', 0: '未标注日期' };

// ---------- 站点名称 ----------
// 使用方可在 config.js 里改成自己单位的叫法，例如「XX市电梯维保单位信用评分」
const SITE_NAME = userCfg.siteName || '电梯维保单位信用评分';
const SITE_SUB = userCfg.siteSub || '记分制 100 − 扣分 + 加分';

// ---------- 只读模式 · 录入密钥 ----------
// 页面默认只读；录入/删除/导入/修改单位信息时需回答密钥问题。
// 校验始终在服务端完成，前端不持有答案。
// ★ 使用方请在同目录 config.js 中修改下面三项（改完重启生效）
const DEF_KEY_QUESTION = '请输入管理密钥';
const DEF_KEY_ANSWER = 'admin';
const DEF_KEY_MSG = '密钥错误，请联系系统管理员';
const KEY_Q = userCfg.keyQuestion || DEF_KEY_QUESTION;
const KEY_A = String(userCfg.keyAnswer || DEF_KEY_ANSWER);
const KEY_MSG = userCfg.keyWrongMsg || DEF_KEY_MSG;
function checkKey(v) {
  return String(v == null ? '' : v).trim().toLowerCase() === KEY_A.trim().toLowerCase();
}

// ---------- 下年度监督检查频次（《评价管理办法》第四章）----------
// 年度检查=年度综合检查/年度检查；监督检查=日常监督检查；证后=持证周期内证后监督检查（按周期摊到年度，仅供参考）
const INSPECT_PLAN = {
  A: { year: '—', sup: '—', post: '—', perYear: null, note: '减少监督检查频次（第十八条）', extra: '优先选聘推荐；换证可自我声明承诺' },
  B: { year: 1, sup: 0, post: 1, perYear: 1, note: '每年 1 次年度综合检查；持证周期内 1 次证后检查', extra: '加强指导帮扶；换证可自我声明承诺' },
  C: { year: 1, sup: 2, post: 2, perYear: 3, note: '每年 1 次年度综合检查 + 至少 2 次监督检查；持证周期内 2 次证后检查', extra: '限期整改并交书面报告；换证不予承诺制；连续两年 C 级约谈主要负责人' },
  D: { year: 1, sup: 4, post: 2, perYear: 5, note: '每年 1 次年度检查 + 每季度 1 次监督检查；持证周期内 2 次证后检查', extra: '约谈主要负责人；换证不予承诺制；区内发证实地监察换证评审，区外发证抄送本部及许可监管部门' }
};
const LEVEL_ORDER = ['D', 'C', 'B', 'A'];

// ---------- 年度评价时间节点（办法第十五条 / 第十七条）----------
const DEADLINES = [
  { month: 1, day: 31, title: '完成上年度信用等级评价并书面告知单位', basis: '第十五条' },
  { month: 3, day: 15, title: '向社会公布上年度信用等级评价结果', basis: '第十七条' }
];

function readDB() {
  const db = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
  // 兼容旧库：缺申请相关字段时自动补齐
  if (!Array.isArray(db.pending_edits)) db.pending_edits = [];
  if (!Array.isArray(db.pending_scores)) db.pending_scores = [];
  if (typeof db.next_edit_id !== 'number') {
    db.next_edit_id = db.pending_edits.reduce((m, x) => Math.max(m, x.id || 0), 0) + 1;
  }
  if (typeof db.next_score_id !== 'number') {
    db.next_score_id = db.pending_scores.reduce((m, x) => Math.max(m, x.id || 0), 0) + 1;
  }
  if (typeof db.next_unit_id !== 'number') {
    db.next_unit_id = db.units.reduce((m, x) => Math.max(m, x.id || 0), 0) + 1;
  }
  return db;
}
function writeDB(db) {
  fs.writeFileSync(DB_FILE, JSON.stringify(db, null, 1), 'utf8');
}

function clauseKey(code) {
  // 扣分/否决项 1..34 排在前，加分1..加分4 排在后
  if (typeof code === 'string' && code.indexOf('加分') === 0) {
    return [1, parseInt(code.replace('加分', ''), 10) || 0];
  }
  return [0, parseInt(code, 10) || 0];
}
function sortByCode(arr) {
  return arr.slice().sort((a, b) => {
    const x = clauseKey(a.code), y = clauseKey(b.code);
    return x[0] - y[0] || x[1] - y[1];
  });
}

/** 由日期得到 {year, q}；无日期/非法返回 {year:0,q:0} */
function quarterOf(date) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(date || '').trim());
  if (!m) return { year: 0, q: 0 };
  const y = parseInt(m[1], 10), mo = parseInt(m[2], 10);
  if (mo < 1 || mo > 12) return { year: 0, q: 0 };
  return { year: y, q: Math.floor((mo - 1) / 3) + 1 };
}
function qlabel(e) {
  const t = quarterOf(e.date);
  return t.q ? `${t.year}年${QUARTER_LABEL[t.q]}` : '未标注日期';
}
/**
 * 否决项（直接评定为D级的情形）记分标准。
 * 依据：桂市监办函〔2026〕94 号《加强智慧监管整治电梯安全"内卷式"竞争工作方案》
 *      （五）深化信用分级监管 3. 明确扣分标准 ——「存在直接评定为 D 级的情形的，一次扣分分值为 50 分」。
 * 每次否决项记 -50 分，同时仍按《评价管理办法》第十一条直接判为 D 级（等级判定优先）。
 */
const VETO_DEDUCT = 50;
/** 记分绝对值 */
function evVal(e) {
  return e.is_veto === '是' ? VETO_DEDUCT : (parseFloat(e.value) || 0);
}
/** 带符号分值（导出使用）：扣分/否决为负，加分为正 */
function evSign(e) {
  if (e.is_veto === '是') return -VETO_DEDUCT;
  return (e.score_type === '加分' ? 1 : -1) * (parseFloat(e.value) || 0);
}
/** 单条记录的扣分值（否决项按 VETO_DEDUCT 计） */
function deductOf(e) {
  return e.is_veto === '是' ? VETO_DEDUCT : (parseFloat(e.value) || 0);
}

function computeSummary(db) {
  const byUnit = new Map();
  db.units.forEach(u => byUnit.set(u.id, { deduct: 0, bonusRaw: 0, veto: false, facts: [] }));
  db.events.forEach(e => {
    const s = byUnit.get(e.unit_id);
    if (!s) return;
    const v = parseFloat(e.value) || 0;
    if (e.is_veto === '是') { s.deduct += VETO_DEDUCT; s.veto = true; }
    else if (e.score_type === '加分') s.bonusRaw += v;
    else s.deduct += v;
    if (e.reason) s.facts.push(`${e.date || '未标注日期'} 第${e.clause_code}条 ${e.is_veto === '是' ? '【否决项】-' + VETO_DEDUCT : (e.score_type === '扣分' ? '-' : '+') + v} ${e.reason}`);
  });

  return db.units.map(u => {
    const s = byUnit.get(u.id) || { deduct: 0, bonusRaw: 0, veto: false, facts: [] };
    let capped = 0;
    Object.keys(BONUS_CAP).forEach(code => {
      const sum = db.events
        .filter(e => e.unit_id === u.id && e.score_type === '加分' && e.clause_code === code)
        .reduce((a, e) => a + (parseFloat(e.value) || 0), 0);
      capped += Math.min(sum, BONUS_CAP[code]);
    });
    const finalScore = 100 - s.deduct + capped;
    let level = 'D';
    if (s.veto) level = 'D';
    else if (finalScore >= 90) level = 'A';
    else if (finalScore >= 80) level = 'B';
    else if (finalScore >= 60) level = 'C';
    return {
      id: u.id, name: u.name, jurisdiction: u.jurisdiction, license_no: u.license_no, note: u.note,
      office: u.office || '', contact: u.contact || '',       phone: u.phone || '',
      scope: u.scope || '', level_no: u.level || '',
      deduct: Math.round(s.deduct * 100) / 100,
      bonus_raw: Math.round(s.bonusRaw * 100) / 100,
      bonus_capped: Math.round(capped * 100) / 100,
      final: Math.round(finalScore * 100) / 100,
      level, veto: s.veto ? '是' : '否', facts: s.facts.join('；'),
      // 下年度检查计划（第四章）
      plan: INSPECT_PLAN[level] || INSPECT_PLAN.A
    };
  }).sort((a, b) => a.final - b.final || a.id - b.id);
}

/** 下年度监督检查计划：按等级分组 */
function inspectPlan(db) {
  const sum = computeSummary(db);
  const groups = LEVEL_ORDER.map(lv => {
    const list = sum.filter(u => u.level === lv);
    const p = INSPECT_PLAN[lv];
    return {
      level: lv, count: list.length, plan: p,
      perYear: p.perYear,
      total: p.perYear == null ? null : p.perYear * list.length,
      units: list.map(u => ({
        id: u.id, name: u.name, jurisdiction: u.jurisdiction, license_no: u.license_no,
        office: u.office, contact: u.contact, phone: u.phone, final: u.final, veto: u.veto
      }))
    };
  });
  const totalTimes = groups.reduce((a, g) => a + (g.total || 0), 0);
  return { groups, totalTimes, deadline: nextDeadline() };
}

/** 距离下一个年度评价节点还有几天 */
function nextDeadline(now) {
  const d = now || new Date();
  const y = d.getFullYear();
  const cands = DEADLINES.map(x => new Date(y, x.month - 1, x.day));
  DEADLINES.forEach((x, i) => cands.push(new Date(y + 1, x.month - 1, x.day)));
  const future = DEADLINES
    .map((x, i) => ({ ...x, date: new Date(y, x.month - 1, x.day) }))
    .concat(DEADLINES.map(x => ({ ...x, date: new Date(y + 1, x.month - 1, x.day) })))
    .filter(x => x.date >= new Date(d.getFullYear(), d.getMonth(), d.getDate()))
    .sort((a, b) => a.date - b.date);
  const next = future[0];
  if (!next) return null;
  const days = Math.ceil((next.date - new Date(d.getFullYear(), d.getMonth(), d.getDate())) / 86400000);
  return {
    title: next.title, basis: next.basis,
    date: `${next.date.getFullYear()}-${String(next.date.getMonth() + 1).padStart(2, '0')}-${String(next.date.getDate()).padStart(2, '0')}`,
    days
  };
}

/** 季度通报：本季度有记分的单位（含否决项）
 *  redact=true 时不返回「违法违规事实」（只读模式对外脱敏） */
function quarterlyOf(db, q, redact) {
  const sum = computeSummary(db);
  const idx = new Map(sum.map(u => [u.id, u]));
  const hit = new Map();
  db.events.forEach(e => {
    const t = quarterOf(e.date);
    // q=null 表示全部；q=0 表示只取未标注日期
    if (q !== null && q !== undefined && t.q !== q) return;
    if (!e.reason && e.score_type !== '扣分' && e.is_veto !== '是') return;
    const u = idx.get(e.unit_id);
    if (!u) return;
    if (!hit.has(u.id)) hit.set(u.id, { facts: [], deduct: 0, bonus: 0, veto: false });
    const r = hit.get(u.id);
    if (e.is_veto === '是') { r.veto = true; r.deduct += VETO_DEDUCT; }
    else if (e.score_type === '加分') r.bonus += (parseFloat(e.value) || 0);
    else r.deduct += (parseFloat(e.value) || 0);
    r.facts.push(`${e.date || '未标注日期'} 第${e.clause_code}条 ${e.is_veto === '是' ? '【否决项】-' + VETO_DEDUCT : (e.score_type === '扣分' ? '-' : '+') + (parseFloat(e.value) || 0)} ${e.reason || ''}`.trim());
  });
  return [...hit.entries()].map(([id, r]) => {
    const u = idx.get(id);
    return {
      id, name: u.name, office: u.office, contact: u.contact, phone: u.phone,
      jurisdiction: u.jurisdiction, license_no: u.license_no,
      facts: redact ? '' : r.facts.join('；'),
      deduct: Math.round(r.deduct * 100) / 100,
      bonus: Math.round(r.bonus * 100) / 100,
      final: u.final, level: u.level, veto: r.veto ? '是' : '否'
    };
  }).sort((a, b) => b.deduct - a.deduct || a.id - b.id);
}

// ---------- 导出用表结构 ----------
function summarySheets(db) {
  const sum = computeSummary(db);
  const stdIdx = new Map(db.standards.map(s => [String(s.code), s]));
  const main = {
    name: '记分汇总',
    cols: [6, 40, 10, 20, 10, 10, 10, 10, 10, 8, 8, 60],
    rows: [[
      '序号', '维保单位名称', '办公点辖区', '许可证号', '许可级别', '备注',
      '累计扣分', '加分(原始)', '加分(封顶后)', '最终得分', '信用等级', '是否否决', '记分明细'
    ]]
  };
  const mainCols = [6, 40, 10, 20, 10, 10, 10, 12, 12, 10, 8, 8, 60];
  main.cols = mainCols;
  sum.forEach((u, i) => {
    const st = (u.veto === '是') ? 3 : 0;
    main.rows.push([
      i + 1, u.name, u.jurisdiction || '', u.license_no || '', u.level_no || '', u.note || '',
      u.deduct, u.bonus_raw, u.bonus_capped,
      { v: u.level === 'D' ? '—' : u.final, s: st }, { v: u.level, s: st },
      u.veto, u.facts || '—'
    ]);
  });

  const det = {
    name: '记分明细',
    cols: [6, 36, 10, 12, 12, 14, 8, 8, 20, 8, 50, 10],
    rows: [['序号', '维保单位名称', '辖区', '日期', '所属季度', '检查类型', '条款编号', '条款分类', '记分项目', '分值', '事由摘要', '录入人']]
  };
  const evs = db.events.slice().sort((a, b) => String(a.date).localeCompare(String(b.date)) || a.id - b.id);
  evs.forEach((e, i) => {
    const s = stdIdx.get(String(e.clause_code)) || {};
    det.rows.push([
      i + 1, e.unit_name, (db.units.find(u => u.id === e.unit_id) || {}).jurisdiction || '',
      e.date || '未标注日期', qlabel(e), e.check_type || '', e.clause_code,
      e.clause_category || s.category || '', s.item || '',
      { v: evSign(e), s: e.score_type === '扣分' ? 3 : 0 },
      (e.is_veto === '是' ? '【否决项·直接定D级】' : '') + (e.reason || ''), e.recorder || ''
    ]);
  });
  return [main, det];
}

function quarterSheet(db, q, titleName, redact) {
  const rows = quarterlyOf(db, q, redact);
  const cols = [6, 36, 30, 10, 10, 10, 10, 8, 60, 10];
  const sheet = {
    name: titleName,
    cols,
    rows: [['序号', '维保单位名称', '办公地点', '办公点辖区', '联系人', '联系电话', '本季扣分', '本季加分', '违法违规事实（记分明细）', '年度信用等级']]
  };
  rows.forEach((u, i) => {
    sheet.rows.push([
      i + 1, u.name, u.office || '', u.jurisdiction || '', u.contact || '', u.phone || '',
      u.deduct, u.bonus, redact ? '—' : (u.facts || '—'), { v: u.level, s: u.veto === '是' ? 3 : 0 }
    ]);
  });
  // 合计行
  const td = rows.reduce((a, u) => a + u.deduct, 0);
  const tb = rows.reduce((a, u) => a + u.bonus, 0);
  sheet.rows.push([{ v: '合计', s: 2 }, { v: `${rows.length} 家单位`, s: 2 }, '', '', '', '',
    { v: Math.round(td * 100) / 100, s: 2 }, { v: Math.round(tb * 100) / 100, s: 2 }, '', '']);
  if (!rows.length) sheet.rows.push(['（本季度暂无记分记录）']);
  return sheet;
}

function quarterDetailSheet(db, q, redact) {
  const stdIdx = new Map(db.standards.map(s => [String(s.code), s]));
  const name = q ? QUARTER_LABEL[q] : '未标注日期';
  const sheet = {
    name: name,
    cols: [6, 36, 10, 30, 10, 10, 14, 8, 20, 8, 50, 10],
    rows: [['序号', '维保单位名称', '办公点辖区', '办公地点', '日期', '条款编号', '条款分类', '记分项目', '检查类型', '分值', '事由摘要', '录入人']]
  };
  const evs = db.events
    .filter(e => quarterOf(e.date).q === q)
    .sort((a, b) => String(a.date).localeCompare(String(b.date)) || a.id - b.id);
  evs.forEach((e, i) => {
    const s = stdIdx.get(String(e.clause_code)) || {};
    sheet.rows.push([
      i + 1, e.unit_name, (db.units.find(u => u.id === e.unit_id) || {}).jurisdiction || '',
      (db.units.find(u => u.id === e.unit_id) || {}).office || '',
      e.date || '未标注日期', e.clause_code, e.clause_category || s.category || '', s.item || '',
      e.check_type || '',
      { v: evSign(e), s: e.score_type === '扣分' ? 3 : 0 },
      redact ? '—' : ((e.is_veto === '是' ? '【否决项·直接定D级】' : '') + (e.reason || '')), e.recorder || ''
    ]);
  });
  const td = evs.filter(e => e.score_type !== '加分').reduce((a, e) => a + evVal(e), 0);
  sheet.rows.push([{ v: '合计', s: 2 }, { v: `${evs.length} 条记分`, s: 2 }, '', '', '', '', '', '', '',
    { v: -Math.round(td * 100) / 100, s: 2 }, '', '']);
  if (!evs.length) sheet.rows.push(['（本季度暂无记分记录）']);
  return sheet;
}

/** 年度监督检查计划表（按等级分组，含每家单位） */
function planSheets(db) {
  const plan = inspectPlan(db);
  const sheets = [];
  const cover = {
    name: '计划汇总',
    cols: [8, 12, 10, 12, 12, 12, 14, 60],
    rows: [['信用等级', '单位家数', '年度检查(次/家)', '监督检查(次/家)', '证后检查(周期内)', '合计(次/家·年)', '全年小计(次)', '频次依据（第四章）']]
  };
  // 单元格样式：0=常规 1=表头 2=粗体 3=红字
  const styleOf = { A: 0, B: 0, C: 2, D: 3 };
  plan.groups.forEach(g => {
    const s = g.plan;
    cover.rows.push([
      { v: g.level + ' 级', s: styleOf[g.level] }, g.count, s.year, s.sup, s.post,
      s.perYear == null ? '减少频次' : s.perYear,
      g.total == null ? '—' : g.total,
      s.note + '；' + s.extra
    ]);
  });
  cover.rows.push([
    { v: '合计', s: 2 }, { v: db.units.length + ' 家', s: 2 }, '', '', '',
    '', { v: plan.totalTimes, s: 2 }, 'A 级按第十八条减少频次，未计入固定次数'
  ]);
  sheets.push(cover);

  // 每个等级一张明细表
  plan.groups.forEach(g => {
    const sh = {
      name: g.level + '级单位',
      cols: [6, 40, 10, 20, 28, 12, 14, 10, 10, 12],
      rows: [['序号', '维保单位名称', '办公点辖区', '许可证号', '办公地点', '联系人', '联系电话', '年度检查', '监督检查', '全年合计']]
    };
    const s = g.plan;
    g.units.forEach((u, i) => {
      sh.rows.push([
        i + 1, u.name, u.jurisdiction || '', u.license_no || '', u.office || '',
        u.contact || '', u.phone || '', s.year, s.sup,
        s.perYear == null ? '减少频次' : s.perYear
      ]);
    });
    if (!g.units.length) sh.rows.push(['（本等级暂无单位）']);
    sheets.push(sh);
  });
  return sheets;
}

function send(res, code, data, type) {
  res.writeHead(code, corsHeaders({ 'Content-Type': type || 'application/json; charset=utf-8' }));
  res.end(typeof data === 'string' || Buffer.isBuffer(data) ? data : JSON.stringify(data));
}

/** 发送文件下载 */
function sendFile(res, filename, buf) {
  res.writeHead(200, corsHeaders({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
    'Content-Length': buf.length
  }));
  res.end(buf);
}
function sendCsv(res, filename, rows) {
  const buf = XLSX.buildCsv(rows);
  res.writeHead(200, corsHeaders({
    'Content-Type': 'text/csv; charset=utf-8',
    'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
    'Content-Length': buf.length
  }));
  res.end(buf);
}

function body(req) {
  return new Promise(r => { let b = ''; req.on('data', c => b += c); req.on('end', () => r(b)); });
}

// ---------- 跨域 ----------
// 资料库云端页面与本服务不同源，需要允许跨域访问（含预检）。
const CORS_ORIGIN = process.env.CORS_ORIGIN || (userCfg && userCfg.corsOrigin) || '*';
function corsHeaders(extra) {
  return Object.assign({
    'Access-Control-Allow-Origin': CORS_ORIGIN,
    'Access-Control-Allow-Methods': 'GET,POST,PATCH,DELETE,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '86400'
  }, extra || {});
}

const server = http.createServer(async (req, res) => {
  const u = new URL(req.url, 'http://x');
  const url = u.pathname;
  const qs = u.searchParams;
  if (req.method === 'OPTIONS') {
    res.writeHead(204, corsHeaders({ 'Content-Length': '0' }));
    return res.end();
  }
  try {
    if (req.method === 'GET' && (url === '/' || url === '/index.html')) {
      return send(res, 200, fs.readFileSync(HTML_FILE), 'text/html; charset=utf-8');
    }
    if (req.method === 'GET' && url === '/api/units') {
      const db = readDB();
      return send(res, 200, db.units.map(x => ({
        id: x.id, name: x.name, jurisdiction: x.jurisdiction, license_no: x.license_no,
        note: x.note, office: x.office || '', contact: x.contact || '', phone: x.phone || '',
        scope: x.scope || '', level: x.level || ''
      })));
    }
    if (req.method === 'GET' && url === '/api/standards') {
      return send(res, 200, sortByCode(readDB().standards));
    }
    if (req.method === 'GET' && url === '/api/events') {
      const db = readDB();
      const q = parseInt(qs.get('q') || '', 10);
      let list = db.events.slice().sort((a, b) => b.id - a.id);
      if (q) list = list.filter(e => quarterOf(e.date).q === q);
      const uid = parseInt(qs.get('unit') || '', 10);
      if (uid) list = list.filter(e => e.unit_id === uid);
      return send(res, 200, list);
    }
    if (req.method === 'GET' && url === '/api/config') {
      return send(res, 200, {
        readOnly: true,
        editKeyQuestion: KEY_Q,
        siteName: SITE_NAME,
        siteSub: SITE_SUB,
        inspectPlan: INSPECT_PLAN,
        deadline: nextDeadline()
      });
    }

    // ---------- 维保单位库：模板下载 + 导入 ----------
    // 模板：CSV（记事本/Excel 都能打开）或 xlsx
    if (req.method === 'GET' && url === '/api/template/units') {
      const rows = [
        ['单位名称', '辖区', '许可证号', '许可级别', '许可范围', '办公地点', '联系人', '联系电话', '备注'],
        ['示例：XX电梯有限公司南宁分公司', '青秀', 'TS3345XXX-2026', 'A', '安装、改造、维修', '南宁市青秀区XX路XX号', '张三', '13800000000', '本市取证']
      ];
      if (qs.get('format') === 'xlsx') {
        return sendFile(res, '维保单位导入模板.xlsx',
          XLSX.build([{ name: '维保单位', cols: [38, 10, 20, 10, 20, 34, 10, 16, 16], rows }]));
      }
      return sendCsv(res, '维保单位导入模板.csv', rows);
    }

    /**
     * 导入维保单位。
     * 入参：{ text: 'CSV文本' } 或 { b64: 'xlsx的base64' }，mode: 'append' | 'replace'，key: 密钥
     * 首次使用（单位库为空）免密钥；库里已有数据时需验证身份。
     * 单位名称相同的行视为同一家，按表内信息更新；其余新增。
     */
    if (req.method === 'POST' && url === '/api/import/units') {
      const db = readDB();
      const p = JSON.parse(await body(req) || '{}');
      const mode = p.mode === 'replace' ? 'replace' : 'append';
      if (db.units.length && !checkKey(p.key)) return send(res, 200, { ok: false, msg: KEY_MSG });

      let rows;
      try {
        if (p.b64) rows = IMP.parseXlsx(Buffer.from(String(p.b64), 'base64'));
        else if (p.text) rows = IMP.parseDelimited(String(p.text));
        else return send(res, 200, { ok: false, msg: '没有收到文件内容，请重新选择文件' });
      } catch (e) {
        return send(res, 200, { ok: false, msg: '文件解析失败：' + (e && e.message || e) + '。建议先下载模板再填写。' });
      }
      if (!rows.length) return send(res, 200, { ok: false, msg: '文件里没有数据' });

      const head = IMP.mapHeader(rows[0]);
      if (head.missing) {
        return send(res, 200, { ok: false, msg: '第一行表头里没有找到【单位名称】列，请先下载模板再填写。' });
      }

      if (mode === 'replace') {
        // 清空重建：单位与其名下的记分、待确认申请一并清掉
        const keepEvents = String(p.keepEvents) === '1';
        db.units = [];
        db.next_unit_id = 1;
        if (!keepEvents) {
          db.events = [];
          db.next_event_id = 1;
          db.pending_scores = [];
          db.pending_edits = [];
          db.next_score_id = 1;
          db.next_edit_id = 1;
        }
      }

      let added = 0, updated = 0, skipped = 0;
      const errors = [];
      const FIELDS = ['name', 'jurisdiction', 'license_no', 'level', 'scope', 'office', 'contact', 'phone', 'note'];
      const idx = new Map();
      db.units.forEach(u => idx.set(String(u.name || '').trim(), u));

      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        const obj = {};
        head.map.forEach((f, c) => { if (f && r[c] !== undefined) obj[f] = String(r[c]).trim(); });
        const name = String(obj.name || '').trim();
        if (!name) { skipped++; continue; }
        if (name.indexOf('示例') === 0 || name.indexOf('示例：') === 0) { skipped++; continue; } // 模板示例行
        if (i > 5000) { errors.push({ row: i + 1, msg: '超出单次导入上限（5000 行），已忽略' }); break; }

        let u = idx.get(name);
        if (u) {
          let changed = false;
          FIELDS.forEach(f => {
            if (f === 'name') return;
            const v = obj[f];
            if (v !== undefined && v !== '' && String(u[f] || '') !== v) { u[f] = v; changed = true; }
          });
          if (changed) updated++; else skipped++;
        } else {
          u = { id: db.next_unit_id++, name };
          FIELDS.forEach(f => { if (f !== 'name') u[f] = obj[f] || ''; });
          u.area = '';
          db.units.push(u);
          idx.set(name, u);
          added++;
        }
      }
      writeDB(db);
      return send(res, 200, {
        ok: true, added, updated, skipped, errors,
        total: db.units.length, mode,
        hasEvents: db.events.length > 0
      });
    }

    // 删除单个单位（需密钥；已有记分记录的先拒绝，避免账对不上）
    if (req.method === 'DELETE' && /^\/api\/units\/\d+$/.test(url)) {
      const id = parseInt(url.split('/').pop(), 10);
      const p = JSON.parse(await body(req) || '{}');
      if (!checkKey(p.key)) return send(res, 200, { ok: false, msg: KEY_MSG });
      const db = readDB();
      const u = db.units.find(x => x.id === id);
      if (!u) return send(res, 200, { ok: false, msg: '单位不存在' });
      const n = db.events.filter(e => e.unit_id === id).length;
      if (n) return send(res, 200, { ok: false, msg: `该单位已有 ${n} 条记分记录，不能删除。如确需删除，请先删除这些记分。` });
      db.units = db.units.filter(x => x.id !== id);
      db.pending_edits = db.pending_edits.filter(x => x.unit_id !== id);
      db.pending_scores = db.pending_scores.filter(x => x.unit_id !== id);
      writeDB(db);
      return send(res, 200, { ok: true, total: db.units.length });
    }
    // 校验录入密钥（答案只在服务端比对）
    if (req.method === 'POST' && url === '/api/unlock') {
      const p = JSON.parse(await body(req) || '{}');
      return send(res, 200, { ok: checkKey(p.answer), msg: checkKey(p.answer) ? '' : KEY_MSG });
    }

    if (req.method === 'POST' && url === '/api/events') {
      const db = readDB();
      const p = JSON.parse(await body(req) || '{}');
      if (!checkKey(p.key)) return send(res, 200, { ok: false, msg: KEY_MSG });
      const unit = db.units.find(x => x.id === +p.unit_id);
      if (!unit) return send(res, 200, { ok: false, msg: '单位不存在' });
      const std = db.standards.find(s => s.code === String(p.clause_code));
      if (!std) return send(res, 200, { ok: false, msg: '条款不存在' });
      const val = parseFloat(p.value);
      if (!(val > 0)) return send(res, 200, { ok: false, msg: '分值需大于0' });
      const ev = {
        id: db.next_event_id++, unit_id: unit.id, unit_name: unit.name,
        date: p.date || '', check_type: p.check_type || '日常监督检查',
        score_type: p.score_type || '扣分', clause_code: std.code,
        clause_category: std.category, is_veto: std.type === '否决项' ? '是' : '否',
        value: std.type === '否决项' ? VETO_DEDUCT : val,
        reason: p.reason || '', recorder: p.recorder || '',
        created_at: new Date().toISOString().slice(0, 19).replace('T', ' ')
      };
      db.events.push(ev);
      writeDB(db);
      return send(res, 200, { ok: true, id: ev.id });
    }
    if (req.method === 'DELETE' && /^\/api\/events\/\d+$/.test(url)) {
      const id = parseInt(url.split('/').pop(), 10);
      if (!checkKey(qs.get('key'))) return send(res, 200, { ok: false, msg: KEY_MSG });
      const db = readDB();
      db.events = db.events.filter(e => e.id !== id);
      writeDB(db);
      return send(res, 200, { ok: true });
    }
    // 修改已有记分记录（需密钥）：单位不可改，其余字段可改
    if ((req.method === 'PUT' || req.method === 'PATCH') && /^\/api\/events\/\d+$/.test(url)) {
      const id = parseInt(url.split('/').pop(), 10);
      const p = JSON.parse(await body(req) || '{}');
      if (!checkKey(p.key)) return send(res, 200, { ok: false, msg: KEY_MSG });
      const db = readDB();
      const ev = db.events.find(e => e.id === id);
      if (!ev) return send(res, 200, { ok: false, msg: '记分记录不存在' });
      const std = db.standards.find(s => s.code === String(p.clause_code));
      if (!std) return send(res, 200, { ok: false, msg: '条款不存在' });
      const val = parseFloat(p.value);
      if (!(val > 0)) return send(res, 200, { ok: false, msg: '分值需大于0' });
      const veto = std.type === '否决项';
      ev.date = String(p.date || '').trim();
      ev.check_type = p.check_type || '日常监督检查';
      ev.score_type = veto ? '扣分' : (p.score_type || '扣分');
      ev.clause_code = std.code;
      ev.clause_category = std.category;
      ev.is_veto = veto ? '是' : '否';
      ev.value = veto ? VETO_DEDUCT : val;
      ev.reason = String(p.reason || '').trim();
      ev.recorder = String(p.recorder || '').trim();
      ev.edited_at = new Date().toISOString().slice(0, 19).replace('T', ' ');
      writeDB(db);
      return send(res, 200, { ok: true, event: ev });
    }
    if (req.method === 'GET' && url === '/api/summary') {
      return send(res, 200, computeSummary(readDB()));
    }
    if (req.method === 'GET' && url === '/api/quarterly') {
      // q 为空=全部；q=0=未标注日期；q=1..4=对应季度
      const raw = qs.get('q');
      const q = (raw === null || raw === '') ? null : (parseInt(raw, 10) || 0);
      // redact=1：只读模式，不返回违法违规事实
      return send(res, 200, quarterlyOf(readDB(), q, qs.get('redact') === '1'));
    }

    // 下年度监督检查计划（第四章频次换算）
    if (req.method === 'GET' && url === '/api/plan') {
      return send(res, 200, inspectPlan(readDB()));
    }
    // ---------- 待确认事项：记分录入申请 + 单位信息修改申请 ----------
    // 只读状态下任何人都可以提交，提交后不生效；由验证过身份的人通过 / 驳回。
    const FIELD_LABELS = { contact: '联系人', phone: '联系电话', office: '办公地址' };
    const stampNow = () => new Date().toISOString().slice(0, 19).replace('T', ' ');
    const pubEdit = x => ({
      id: x.id, unit_id: x.unit_id, unit_name: x.unit_name, field: x.field, field_label: x.field_label,
      old_value: x.old_value, new_value: x.new_value, requester: x.requester, created_at: x.created_at
    });
    const pubScore = x => ({
      id: x.id, unit_id: x.unit_id, unit_name: x.unit_name, date: x.date, check_type: x.check_type,
      score_type: x.score_type, clause_code: x.clause_code, clause_item: x.clause_item,
      value: x.value, is_veto: x.is_veto, reason: x.reason, recorder: x.recorder,
      requester: x.requester, created_at: x.created_at
    });

    if (req.method === 'GET' && url === '/api/pending') {
      const db = readDB();
      const scores = db.pending_scores.filter(x => x.status === 'pending').map(pubScore);
      const edits = db.pending_edits.filter(x => x.status === 'pending').map(pubEdit);
      return send(res, 200, { scores, edits, count: scores.length + edits.length });
    }

    // 提交一条记分录入申请（无需密钥）
    if (req.method === 'POST' && url === '/api/pending/score') {
      const db = readDB();
      const p = JSON.parse(await body(req) || '{}');
      const unit = db.units.find(x => x.id === +p.unit_id);
      if (!unit) return send(res, 200, { ok: false, msg: '请选择维保单位' });
      const std = db.standards.find(s => s.code === String(p.clause_code));
      if (!std) return send(res, 200, { ok: false, msg: '条款不存在' });
      const val = parseFloat(p.value);
      if (!(val > 0)) return send(res, 200, { ok: false, msg: '分值需大于0' });
      if (!String(p.date || '').trim()) return send(res, 200, { ok: false, msg: '请填写日期' });
      const rec = String(p.recorder || '').trim() || '匿名';
      const dup = db.pending_scores.find(x => x.status === 'pending' && x.unit_id === unit.id
        && String(x.clause_code) === String(std.code) && String(x.date) === String(p.date)
        && Math.abs(x.value - val) < 1e-9);
      if (dup) return send(res, 200, { ok: false, msg: '同一单位同一天的同一条款已有申请，等待确认' });
      const item = {
        id: db.next_score_id++, unit_id: unit.id, unit_name: unit.name,
        date: String(p.date || '').trim(), check_type: p.check_type || '日常监督检查',
        score_type: p.score_type || (std.type === '加分' ? '加分' : '扣分'),
        clause_code: std.code, clause_category: std.category, clause_item: std.item,
        is_veto: std.type === '否决项' ? '是' : '否',
        value: std.type === '否决项' ? VETO_DEDUCT : val,
        reason: String(p.reason || '').trim(), recorder: rec, requester: rec,
        status: 'pending', created_at: stampNow(), decided_at: '', decided_by: ''
      };
      db.pending_scores.push(item);
      writeDB(db);
      return send(res, 200, { ok: true, item: pubScore(item) });
    }

    // 提交单位信息修改申请（无需密钥）
    if (req.method === 'POST' && url === '/api/pending/edit') {
      const db = readDB();
      const p = JSON.parse(await body(req) || '{}');
      const field = FIELD_LABELS[p.field] ? p.field : null;
      if (!field) return send(res, 200, { ok: false, msg: '字段不正确' });
      const unit = db.units.find(x => x.id === +p.unit_id);
      if (!unit) return send(res, 200, { ok: false, msg: '单位不存在' });
      const value = String(p.value || '').trim();
      if (!value) return send(res, 200, { ok: false, msg: '内容不能为空' });
      if (String(unit[field] || '').trim() === value) {
        return send(res, 200, { ok: false, msg: '与现登记信息一致，无需修改' });
      }
      const dup = db.pending_edits.find(x => x.status === 'pending' && x.unit_id === unit.id && x.field === field && x.new_value === value);
      if (dup) return send(res, 200, { ok: false, msg: '已有相同内容的申请，正在等待确认' });
      const item = {
        id: db.next_edit_id++, unit_id: unit.id, unit_name: unit.name,
        field, field_label: FIELD_LABELS[field],
        old_value: String(unit[field] || ''), new_value: value,
        requester: String(p.requester || '').trim() || '匿名',
        status: 'pending', created_at: stampNow(), decided_at: '', decided_by: ''
      };
      db.pending_edits.push(item);
      writeDB(db);
      return send(res, 200, { ok: true, item: pubEdit(item) });
    }

    // 通过 / 驳回（需密钥）：kind = score | edit
    if (req.method === 'POST' && url === '/api/pending/decide') {
      const p = JSON.parse(await body(req) || '{}');
      if (!checkKey(p.key)) return send(res, 200, { ok: false, msg: KEY_MSG });
      const kind = p.kind === 'edit' ? 'edit' : 'score';
      const db = readDB();
      const arr = kind === 'edit' ? db.pending_edits : db.pending_scores;
      const item = arr.find(x => x.id === +p.id);
      if (!item) return send(res, 200, { ok: false, msg: '申请不存在' });
      if (item.status !== 'pending') return send(res, 200, { ok: false, msg: '该申请已处理过了' });
      const unit = db.units.find(x => x.id === item.unit_id);
      if (!unit) return send(res, 200, { ok: false, msg: '单位不存在' });
      item.decided_at = stampNow();
      if (p.action === 'reject') {
        item.status = 'rejected'; writeDB(db);
        return send(res, 200, { ok: true, status: 'rejected' });
      }
      if (kind === 'edit') {
        unit[item.field] = item.new_value;
        item.status = 'approved';
        writeDB(db);
        return send(res, 200, {
          ok: true, status: 'approved',
          unit: { id: unit.id, contact: unit.contact, phone: unit.phone, office: unit.office }
        });
      }
      const std = db.standards.find(s => s.code === String(item.clause_code));
      if (!std) return send(res, 200, { ok: false, msg: '条款不存在，无法计入' });
      const ev = {
        id: db.next_event_id++, unit_id: unit.id, unit_name: unit.name,
        date: item.date, check_type: item.check_type,
        score_type: item.score_type, clause_code: std.code,
        clause_category: std.category, is_veto: std.type === '否决项' ? '是' : '否',
        value: std.type === '否决项' ? VETO_DEDUCT : (parseFloat(item.value) || 0),
        reason: item.reason || '', recorder: item.recorder || '',
        created_at: stampNow()
      };
      db.events.push(ev);
      item.status = 'approved';
      writeDB(db);
      return send(res, 200, { ok: true, status: 'approved', event: ev });
    }

    // 修改单位联系信息（联系人 / 电话 / 地址），需密钥
    if (req.method === 'PATCH' && /^\/api\/units\/\d+$/.test(url)) {
      const id = parseInt(url.split('/').pop(), 10);
      const p = JSON.parse(await body(req) || '{}');
      if (!checkKey(p.key)) return send(res, 200, { ok: false, msg: KEY_MSG });
      const db = readDB();
      const u = db.units.find(x => x.id === id);
      if (!u) return send(res, 200, { ok: false, msg: '单位不存在' });
      if (p.contact !== undefined) u.contact = String(p.contact || '').trim();
      if (p.phone !== undefined) u.phone = String(p.phone || '').trim();
      if (p.office !== undefined) u.office = String(p.office || '').trim();
      writeDB(db);
      return send(res, 200, { ok: true, unit: { id: u.id, contact: u.contact, phone: u.phone, office: u.office } });
    }

    // ---------- 导出 ----------
    if (req.method === 'GET' && url === '/api/export/plan') {
      const db = readDB();
      const stamp = new Date().toISOString().slice(0, 10);
      return sendFile(res, `电梯维保单位年度监督检查计划_${stamp}.xlsx`, XLSX.build(planSheets(db)));
    }
    if (req.method === 'GET' && url === '/api/export/summary') {
      const db = readDB();
      const stamp = new Date().toISOString().slice(0, 10);
      if (qs.get('format') === 'csv') {
        const [main] = summarySheets(db);
        return sendCsv(res, `电梯维保单位信用记分汇总_${stamp}.csv`, main.rows);
      }
      return sendFile(res, `电梯维保单位信用记分汇总_${stamp}.xlsx`, XLSX.build(summarySheets(db)));
    }
    if (req.method === 'GET' && url === '/api/export/quarterly') {
      const db = readDB();
      const raw = qs.get('q');
      const hasQ = !(raw === null || raw === '');
      const q = hasQ ? (parseInt(raw, 10) || 0) : 0;
      const stamp = new Date().toISOString().slice(0, 10);
      const rd = qs.get('redact') === '1';   // 只读模式导出：隐去违法违规事实
      const sheets = hasQ
        ? [quarterSheet(db, q, QUARTER_LABEL[q] + '通报', rd)]
        : [1, 2, 3, 4].map(n => quarterSheet(db, n, QUARTER_LABEL[n] + '通报', rd))
          .concat([quarterSheet(db, 0, '未标注日期', rd)]);
      if (qs.get('format') === 'csv') {
        return sendCsv(res, `季度通报_${hasQ ? QUARTER_LABEL[q] : '全年度'}_${stamp}.csv`, sheets[0].rows);
      }
      return sendFile(res, `电梯维保单位季度通报_${hasQ ? QUARTER_LABEL[q] : '全年度'}_${stamp}.xlsx`, XLSX.build(sheets));
    }
    if (req.method === 'GET' && url === '/api/export/quarters') {
      const db = readDB();
      const stamp = new Date().toISOString().slice(0, 10);
      const raw = qs.get('q');
      const hasQ = !(raw === null || raw === '');
      const q = hasQ ? (parseInt(raw, 10) || 0) : 0;
      const rd = qs.get('redact') === '1';   // 只读模式导出：隐去事由摘要
      const sheets = hasQ
        ? [quarterDetailSheet(db, q, rd)]
        : [1, 2, 3].map(n => quarterDetailSheet(db, n, rd)).concat([quarterDetailSheet(db, 0, rd)]);
      if (qs.get('format') === 'csv') {
        return sendCsv(res, `${hasQ ? QUARTER_LABEL[q] : '前三季度'}记分情况_${stamp}.csv`, sheets[0].rows);
      }
      return sendFile(res, `电梯维保单位${hasQ ? QUARTER_LABEL[q] : '前三季度'}记分情况_${stamp}.xlsx`, XLSX.build(sheets));
    }
    if (req.method === 'GET' && url === '/api/stats') {
      const db = readDB();
      const byQ = { 1: 0, 2: 0, 3: 0, 4: 0, 0: 0 };
      db.events.forEach(e => { byQ[quarterOf(e.date).q]++; });
      return send(res, 200, { events: db.events.length, units: db.units.length, byQuarter: byQ });
    }

    return send(res, 404, { ok: false, msg: 'not found' });
  } catch (err) {
    return send(res, 500, { ok: false, msg: String(err && err.message || err) });
  }
});

server.listen(PORT, process.env.HOST || '0.0.0.0', () => console.log('listening on ' + PORT));
