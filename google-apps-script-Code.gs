// ═══════════════════════════════════════════════════════════════
// DLIG Command Centre — Google Apps Script v2
// Task Board: 双向同步原版 tab
// Marketing / GDC: 存进 CommandCentre tab 做存档
// ═══════════════════════════════════════════════════════════════

const SHEET_IDS = {
  tasks: '1A3g_WPDU-R4zU8gj8gGHu885z5bk4lTBBRnCfDsstDI',
  mkt:   '1F0Ss1MuAwVRkVfWch2wepx2SvMXG3BoIBfVN0v9ye9M',
  gdc:   '1Gc0rO-gx_CBSZ60fFHvmMeTsq7dClsCU4WiSN5ii-2M',
  pay:   '1zGy3rV0bv2dERFRWoRFhRLGj68oJ3RtAj-kiqi71DXk'  // 分钱计算 Sheet
};

const TASK_TAB      = 'Task Board';
const TASK_HDR_ROW  = 4;   // 第4行是 header
const TASK_DATA_ROW = 5;   // 第5行起是数据

// Marketing Content tab (source tab name in mkt spreadsheet)
const MKT_TAB   = 'Marketing Content';
// Sales Dashboard tab (source tab in mkt spreadsheet — website reads, user edits)
const SALES_TAB = 'Sales Dashboard';

const MKT_HEADERS   = ['id','title','date','platform','cat','person','status','doneDate','actual','eff'];
const GDC_HEADERS   = ['id','task','date','dl','cat','who','pic','status','pay'];
const SALES_HEADERS = ['month','target','actual','xd','ylyd','exp','book','note'];

// ─── 工具函数 ────────────────────────────────────────────────

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function fmtDate(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).trim();
}

function calcEff(deadline, doneDate) {
  if (!deadline || !doneDate) return '';
  try {
    const dl   = new Date(deadline);
    const done = new Date(doneDate);
    if (isNaN(dl) || isNaN(done)) return '';
    const diff = Math.floor((done - dl) / 86400000);
    if (diff <= -2) return 1.5;
    if (diff <= 0)  return 1;
    if (diff <= 3)  return 0.5;
    return 0;
  } catch(e) { return ''; }
}

function typeToLabel(type) {
  if (type === 'fee') return 'Task Fee 类（按件算钱）';
  return '贡献时间类（分红）';
}

function labelToType(label) {
  return String(label || '').includes('Fee') ? 'fee' : 'contrib';
}

function normPriority(raw) {
  const s = String(raw || '').trim();
  if (s.includes('紧急') && !s.startsWith('🔴')) return '🔴 紧急';
  if (s.includes('本周') && !s.startsWith('🟡')) return '🟡 本周';
  if (s.includes('本月') && !s.startsWith('⚡'))  return '⚡ 本月';
  if (s.includes('长期') && !s.startsWith('⭐'))  return '⭐ 长期';
  return s;
}

// ─── TASK BOARD 读取 ─────────────────────────────────────────

function readTaskBoard() {
  const tab = SpreadsheetApp.openById(SHEET_IDS.tasks).getSheetByName(TASK_TAB);
  if (!tab) return { error: 'Task Board tab not found' };

  const all = tab.getDataRange().getValues();
  if (all.length < TASK_HDR_ROW) return [];

  // 建立 header → 列索引 映射
  const hdrs = all[TASK_HDR_ROW - 1];
  const h = {};
  hdrs.forEach((name, i) => { h[String(name).trim()] = i; });

  const results = [];
  all.slice(TASK_DATA_ROW - 1).forEach((row, idx) => {
    const task = String(row[h['任务内容'] ?? 0] || '').trim();
    if (!task) return;

    const deadline = fmtDate(row[h['截止日期']]);
    const doneDate = fmtDate(row[h['完成日期']]);
    const actual   = String(row[h['用时']]  ?? '').trim();
    // 效率从 Sheet 读取（如果有），否则留空（网站会动态计算显示）
    const eff      = String(row[h['效率']]  ?? '').trim();

    // 没有 status 列 → 根据完成日期推断
    const status = doneDate ? '已完成' : '待启动';

    results.push({
      id:       idx + 1,
      task,
      person:   String(row[h['负责人']] ?? '').trim(),
      cat:      String(row[h['类别']]   ?? '').trim(),
      type:     labelToType(row[h['类型']]),
      priority: normPriority(row[h['优先级']]),
      deadline,
      doneDate,
      actual,
      eff,
      status,
      desc:     '',   // desc 是网站专用，不从 Sheet 读取
      est:      0
    });
  });

  return results;
}

// ─── TASK BOARD 写入 ─────────────────────────────────────────

function writeTaskBoard(tasks) {
  const tab = SpreadsheetApp.openById(SHEET_IDS.tasks).getSheetByName(TASK_TAB);
  if (!tab) return { error: 'Task Board tab not found' };

  const all  = tab.getDataRange().getValues();
  const hdrs = all[TASK_HDR_ROW - 1];
  const h    = {};
  hdrs.forEach((name, i) => { h[String(name).trim()] = i; });
  const numCols = hdrs.length;

  // 保留 参考/SOP 和 备注（按任务内容匹配）
  const preserved = {};
  all.slice(TASK_DATA_ROW - 1).forEach(row => {
    const key = String(row[h['任务内容'] ?? 0] || '').trim();
    if (key) preserved[key] = {
      sop:  h['参考/SOP'] !== undefined ? row[h['参考/SOP']] : '',
      note: h['备注']     !== undefined ? row[h['备注']]     : ''
    };
  });

  // 清除数据行（保留第1-3行标题 & 第4行 header）
  const lastRow = tab.getLastRow();
  if (lastRow >= TASK_DATA_ROW) {
    tab.getRange(TASK_DATA_ROW, 1, lastRow - TASK_DATA_ROW + 1, numCols).clearContent();
  }

  if (!tasks.length) return { ok: true, count: 0 };

  const rows = tasks.map(item => {
    const key  = String(item.task || '').trim();
    const pres = preserved[key] || { sop: '', note: '' };
    const eff  = calcEff(item.deadline, item.doneDate);

    const row = new Array(numCols).fill('');
    if (h['任务内容']  !== undefined) row[h['任务内容']]  = item.task     || '';
    if (h['参考/SOP'] !== undefined) row[h['参考/SOP']] = pres.sop;
    if (h['负责人']   !== undefined) row[h['负责人']]   = item.person   || '';
    if (h['类别']     !== undefined) row[h['类别']]     = item.cat      || '';
    if (h['类型']     !== undefined) row[h['类型']]     = typeToLabel(item.type);
    if (h['优先级']   !== undefined) row[h['优先级']]   = item.priority || '';
    if (h['截止日期'] !== undefined) row[h['截止日期']] = item.deadline || '';
    if (h['完成日期'] !== undefined) row[h['完成日期']] = item.doneDate || '';
    if (h['用时']     !== undefined) row[h['用时']]     = item.actual   || '';
    if (h['效率']     !== undefined) row[h['效率']]     = eff !== ''    ? eff : '';
    if (h['备注']     !== undefined) row[h['备注']]     = pres.note;
    return row;
  });

  tab.getRange(TASK_DATA_ROW, 1, rows.length, numCols).setValues(rows);
  return { ok: true, count: rows.length };
}

// ─── COMMANDCENTRE tab（Marketing / GDC）────────────────────

function getCCTab(sheetId, headers) {
  const ss = SpreadsheetApp.openById(sheetId);
  let tab  = ss.getSheetByName('CommandCentre');
  if (!tab) {
    tab = ss.insertSheet('CommandCentre');
    tab.getRange(1, 1, 1, headers.length)
       .setValues([headers]).setFontWeight('bold').setBackground('#fdf8f2');
  }
  return tab;
}

function readCC(sheetId, headers) {
  const tab  = getCCTab(sheetId, headers);
  const data = tab.getDataRange().getValues();
  if (data.length <= 1) return [];
  const hdrs = data[0];
  return data.slice(1).filter(r => r[0] !== '').map(row => {
    const obj = {};
    hdrs.forEach((h, i) => { obj[h] = typeof row[i] === 'number' ? row[i] : String(row[i] ?? ''); });
    return obj;
  });
}

function writeCC(sheetId, headers, items) {
  const tab     = getCCTab(sheetId, headers);
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (!items.length) return { ok: true, count: 0 };
  const rows = items.map(item => headers.map(h => item[h] ?? ''));
  tab.getRange(2, 1, rows.length, headers.length).setValues(rows);
  return { ok: true, count: rows.length };
}

// ─── SALES DASHBOARD ────────────────────────────────────────

function readSalesDashboard() {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.mkt);
  const tab = ss.getSheetByName(SALES_TAB);
  if (!tab) return { error: 'Sales Dashboard tab not found' };

  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];

  // ── 找真正的 header 行（含"月份"的那行）────────────────────
  let hdrIdx = -1;
  for (let i = 0; i < Math.min(data.length, 15); i++) {
    const first = String(data[i][0] || '').trim();
    if (first === '月份' || first === 'Month') { hdrIdx = i; break; }
  }
  if (hdrIdx < 0) return { error: 'Header row (月份) not found in first 15 rows' };

  const hdrs = data[hdrIdx].map(h => String(h).trim());
  const h = {};
  hdrs.forEach((name, i) => { if (name) h[name] = i; });

  const col = (...names) => {
    for (const n of names) { if (h[n] !== undefined) return h[n]; }
    return -1;
  };
  // 把 "RM10,000" / 10000 都转成数字
  const num = v => parseFloat(String(v||'').replace(/[^0-9.\-]/g,'')) || 0;

  // 只保留 "Mon YYYY" 格式的月份行（Jan 2026…），过滤标题/合计/产品拆解行
  return data.slice(hdrIdx + 1).filter(row => {
    const f = String(row[0]||'').trim();
    return /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$/.test(f);
  }).map(row => {
    const get = (...names) => { const i = col(...names); return i >= 0 ? row[i] : ''; };
    return {
      month:  String(get('月份','Month') || '').trim(),
      target: num(get('Target (RM)','Target','目标')),
      actual: num(get('Actual (RM)','Actual','实际')),
      xd:     num(get('心动觉察','XD')),
      ylyd:   num(get('YLYD System','YLYD')),
      exp:    num(get('YLYD 体验包','体验包')),
      book:   num(get('设计册','Book')),
      note:   String(get('备注','Note') || '').trim()
    };
  });
}

function writeSalesDashboard(items) {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.mkt);
  const tab = ss.getSheetByName(SALES_TAB);
  if (!tab) return { error: 'Sales Dashboard tab not found' };

  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2, 1, lastRow - 1, SALES_HEADERS.length).clearContent();
  if (!items.length) return { ok: true, count: 0 };
  const rows = items.map(item => SALES_HEADERS.map(k => item[k] ?? ''));
  tab.getRange(2, 1, rows.length, SALES_HEADERS.length).setValues(rows);
  return { ok: true, count: rows.length };
}

// ─── 分钱计算 — 写入汇总 Sheet ──────────────────────────────

const PAY_HEADERS = ['来源','id','任务内容','负责人','类别','类型','截止日期','完成日期','用时','效率','分钱系数','应付金额'];
const COEF_MAP   = {Sales:3,Marketing:2,System:2,GDC:2,'GDC job':2,Operation:1,Admin:1};

function writePaySheet(tasks, mkt, gdc) {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.pay);
  // 找或创建 "分钱汇总" tab
  let tab = ss.getSheetByName('分钱汇总');
  if (!tab) {
    tab = ss.insertSheet('分钱汇总');
    tab.getRange(1,1,1,PAY_HEADERS.length).setValues([PAY_HEADERS]).setFontWeight('bold').setBackground('#fdf8f2');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2,1,lastRow-1,PAY_HEADERS.length).clearContent();

  const rows = [];

  // Task Board 任务
  (tasks||[]).forEach(t => {
    if (!t.person || t.person==='All' || t.person==='全员') return;
    const eff   = t.eff || calcEff(t.deadline, t.doneDate) || '';
    const coef  = COEF_MAP[t.cat] || 1;
    const hours = parseFloat(t.actual) || 0;
    const pay   = t.type==='fee' ? '' : (hours * coef * (parseFloat(eff)||1)).toFixed(1);
    rows.push(['TaskBoard', t.id||'', t.task||'', t.person||'', t.cat||'', t.type==='fee'?'Fee':'贡献', t.deadline||'', t.doneDate||'', hours||'', eff, coef, pay]);
  });

  // GDC 任务
  (gdc||[]).forEach(t => {
    if (!t.doneDate) return; // 只记录已完成
    const eff = calcEff(t.dl, t.doneDate) || '';
    rows.push(['GDC', t.id||'', t.task||'', t.pic||'', 'GDC', 'Fee', t.dl||'', t.doneDate||'', '', eff, 2, '']);
  });

  // Marketing 内容
  (mkt||[]).forEach(m => {
    if (!m.doneDate) return;
    const eff = calcEff(m.dl, m.doneDate) || '';
    rows.push(['Marketing', m.id||'', m.title||'', m.person||'', 'Marketing', 'Fee', m.dl||'', m.doneDate||'', '', eff, 2, '']);
  });

  if (rows.length > 0) tab.getRange(2,1,rows.length,PAY_HEADERS.length).setValues(rows);
  return { ok: true, count: rows.length };
}

// ─── doGet ───────────────────────────────────────────────────

function doGet(e) {
  const type = (e.parameter && e.parameter.sheet) || '';
  try {
    if (type === 'tasks') return respond(readTaskBoard());
    if (type === 'mkt')   return respond(readCC(SHEET_IDS.mkt, MKT_HEADERS));
    if (type === 'gdc')   return respond(readCC(SHEET_IDS.gdc, GDC_HEADERS));
    if (type === 'sales') return respond(readSalesDashboard());
    return respond({ error: 'Unknown sheet: ' + type });
  } catch(err) {
    return respond({ error: err.message });
  }
}

// ─── doPost ──────────────────────────────────────────────────

function doPost(e) {
  const type = (e.parameter && e.parameter.sheet) || '';
  try {
    const data = JSON.parse(e.postData.contents);
    if (!Array.isArray(data)) return respond({ error: 'Expected JSON array' });
    if (type === 'tasks') return respond(writeTaskBoard(data));
    if (type === 'mkt')   return respond(writeCC(SHEET_IDS.mkt, MKT_HEADERS, data));
    if (type === 'gdc')   return respond(writeCC(SHEET_IDS.gdc, GDC_HEADERS, data));
    if (type === 'sales') return respond(writeSalesDashboard(data));
    if (type === 'pay') {
      // data = { tasks:[...], gdc:[...], mkt:[...] }
      return respond(writePaySheet(data.tasks, data.mkt, data.gdc));
    }
    return respond({ error: 'Unknown sheet: ' + type });
  } catch(err) {
    return respond({ error: err.message });
  }
}
