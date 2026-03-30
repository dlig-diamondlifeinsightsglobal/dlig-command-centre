// ═══════════════════════════════════════════════════════════════
// DLIG Command Centre — Google Apps Script v3
// Task Board:        双向同步 Task Board tab（第1行 header）
// Marketing Content: 直接读写 Sales & Marketing 的 Marketing Content tab
// GDC:               直接读写 GDC spreadsheet 的 GDC_JOB_TAB tab
// Events & Rotation: 读 Task Board spreadsheet 里的 "Events 活动" & "Rotation 轮值" tab
// ═══════════════════════════════════════════════════════════════

const SHEET_IDS = {
  tasks: '1A3g_WPDU-R4zU8gj8gGHu885z5bk4lTBBRnCfDsstDI',
  mkt:   '1F0Ss1MuAwVRkVfWch2wepx2SvMXG3BoIBfVN0v9ye9M',
  gdc:   '1Gc0rO-gx_CBSZ60fFHvmMeTsq7dClsCU4WiSN5ii-2M',
  admin: '1zGy3rV0bv2dERFRWoRFhRLGj68oJ3RtAj-kiqi71DXk'
};

const TASK_TAB      = 'Task Board';
const TASK_HDR_ROW  = 1;
const TASK_DATA_ROW = 2;

const MKT_TAB       = 'Marketing Content';  // tab 名在 Sales & Marketing spreadsheet
const GDC_JOB_TAB   = 'Content arrangement';  // GDC spreadsheet 的实际 tab 名
const SALES_TAB     = 'Sales Dashboard';
const EVENTS_TAB    = 'Events 活动';         // 在 Task Board spreadsheet 里新建这个 tab
const ROTATION_TAB  = 'Rotation 轮值';       // 在 Task Board spreadsheet 里新建这个 tab

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

// ─── 灵活的列查找工具 ─────────────────────────────────────────

function makeFlexFinder(hdrs) {
  const lower = hdrs.map(h => String(h||'').trim().toLowerCase());
  return function findCol(...names) {
    for (const n of names) {
      const i = lower.indexOf(n.toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  };
}

// ─── TASK BOARD 读取 ─────────────────────────────────────────

function readTaskBoard() {
  const tab = SpreadsheetApp.openById(SHEET_IDS.tasks).getSheetByName(TASK_TAB);
  if (!tab) return { error: 'Task Board tab not found' };

  const all = tab.getDataRange().getValues();
  if (all.length < TASK_HDR_ROW) return [];

  const hdrs = all[TASK_HDR_ROW - 1];
  const h = {};
  hdrs.forEach((name, i) => { h[String(name).trim()] = i; });

  const results = [];
  all.slice(TASK_DATA_ROW - 1).forEach((row, idx) => {
    const task = String(row[h['任务内容'] ?? 0] || '').trim();
    if (!task) return;

    const deadline   = fmtDate(row[h['截止日期']]);
    const doneDateRaw= fmtDate(row[h['完成日期']]);
    const actual     = String(row[h['用时']]  ?? '').trim();
    const eff        = String(row[h['效率']]  ?? '').trim();
    // 只有完成日期 AND 用时都有填写，才算真正完成
    // 避免 Sheet 里「完成日期」列有默认值但实际上还没完成的误判
    const isRealDone = doneDateRaw && actual && actual !== '0';
    const doneDate   = isRealDone ? doneDateRaw : '';
    const status     = isRealDone ? '已完成' : '待启动';

    results.push({
      id:       idx + 100,
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
      desc:     '',
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

  const preserved = {};
  all.slice(TASK_DATA_ROW - 1).forEach(row => {
    const key = String(row[h['任务内容'] ?? 0] || '').trim();
    if (key) preserved[key] = {
      sop:  h['参考/SOP'] !== undefined ? row[h['参考/SOP']] : '',
      note: h['备注']     !== undefined ? row[h['备注']]     : ''
    };
  });

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

// ─── MARKETING CONTENT 读取（直接读原版 tab）────────────────

function readMarketingContent() {
  const ss = SpreadsheetApp.openById(SHEET_IDS.mkt);
  const tab = ss.getSheetByName(MKT_TAB);
  if (!tab) return { error: 'Marketing Content tab not found in S&M spreadsheet' };

  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];

  const hdrs = data[0];
  const find = makeFlexFinder(hdrs);

  const col = {
    id:       find('id'),
    title:    find('内容标题','title','标题','发布内容','内容','content'),
    date:     find('发布日期','date','日期'),
    platform: find('平台','platform'),
    type:     find('类型','type','内容类型','content type'),
    cat:      find('类别','cat','category'),
    person:   find('负责人','person','who','负责'),
    dl:       find('截止日期','deadline','dl','due'),
    doneDate: find('完成日期','done date','donedate'),
    actual:   find('用时','actual','hours'),
    eff:      find('效率','eff','efficiency'),
    status:   find('状态','status')
  };

  const g  = (row, c) => c >= 0 ? String(row[c] ?? '').trim() : '';
  const fd = (row, c) => c >= 0 ? fmtDate(row[c]) : '';

  return data.slice(1)
    .filter(r => g(r, col.title) || fd(r, col.date))
    .map((row, idx) => ({
      id:       g(row, col.id)   || ('mc' + (idx + 100)),
      title:    g(row, col.title),
      date:     fd(row, col.date)     || g(row, col.date),
      platform: g(row, col.platform),
      type:     g(row, col.type)      || g(row, col.cat),
      cat:      g(row, col.cat)       || g(row, col.type),
      person:   g(row, col.person),
      dl:       fd(row, col.dl)       || g(row, col.dl),
      doneDate: fd(row, col.doneDate) || g(row, col.doneDate),
      actual:   g(row, col.actual),
      eff:      g(row, col.eff),
      status:   g(row, col.status)    || '待发布'
    }));
}

// ─── MARKETING CONTENT 写入（直接写原版 tab）────────────────

function writeMarketingContent(items) {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.mkt);
  const tab = ss.getSheetByName(MKT_TAB);
  if (!tab) return { error: 'Marketing Content tab not found' };

  const hdrs    = tab.getRange(1, 1, 1, tab.getLastColumn()).getValues()[0];
  const find    = makeFlexFinder(hdrs);
  const numCols = hdrs.length;

  const colMap = {
    id:       find('id'),
    title:    find('内容标题','title','标题','发布内容','内容','content'),
    date:     find('发布日期','date','日期'),
    platform: find('平台','platform'),
    type:     find('类型','type','内容类型'),
    cat:      find('类别','cat'),
    person:   find('负责人','person','who'),
    dl:       find('截止日期','deadline','dl'),
    doneDate: find('完成日期','donedate'),
    actual:   find('用时','actual'),
    eff:      find('效率','eff'),
    status:   find('状态','status')
  };

  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2, 1, lastRow - 1, numCols).clearContent();
  if (!items.length) return { ok: true, count: 0 };

  const rows = items.map(item => {
    const row = new Array(numCols).fill('');
    Object.entries(colMap).forEach(([field, idx]) => {
      if (idx >= 0 && item[field] !== undefined) row[idx] = item[field] ?? '';
    });
    return row;
  });

  tab.getRange(2, 1, rows.length, numCols).setValues(rows);
  return { ok: true, count: rows.length };
}

// ─── GDC MARKETING JOB 读取（直接读原版 tab）───────────────

function readGDCJobs() {
  const ss = SpreadsheetApp.openById(SHEET_IDS.gdc);
  // 尝试多个可能的 tab 名
  let tab = ss.getSheetByName(GDC_JOB_TAB);
  if (!tab) {
    for (const n of ['GDC Job','Marketing Job','Jobs','GDC','Sheet1','工作表1']) {
      tab = ss.getSheetByName(n);
      if (tab) break;
    }
  }
  if (!tab) tab = ss.getSheets()[0]; // 拿第一个 tab 作 fallback
  if (!tab) return { error: 'No tab found in GDC spreadsheet' };

  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];

  const hdrs = data[0];
  const find = makeFlexFinder(hdrs);

  const col = {
    id:     find('id'),
    task:   find('任务内容','task','内容','job','工作'),
    date:   find('发布日期','date','日期'),
    dl:     find('截止日期','deadline','dl','due'),
    cat:    find('类别','cat','type','类型','category'),
    who:    find('负责人','who','person','负责 gdc','负责'),
    pic:    find('pic','截图','screenshot'),
    status: find('状态','status'),
    pay:    find('pay','费用','amount','金额')
  };

  const g  = (row, c) => c >= 0 ? String(row[c] ?? '').trim() : '';
  const fd = (row, c) => c >= 0 ? fmtDate(row[c]) : '';

  return data.slice(1)
    .filter(r => g(r, col.task))
    .map((row, idx) => ({
      id:     g(row, col.id) || ('gdc' + (idx + 100)),
      task:   g(row, col.task),
      date:   fd(row, col.date)     || g(row, col.date),
      dl:     fd(row, col.dl)       || g(row, col.dl),
      cat:    g(row, col.cat),
      who:    g(row, col.who),
      pic:    g(row, col.pic),
      status: g(row, col.status)    || '待发布',
      pay:    g(row, col.pay)
    }));
}

// ─── GDC MARKETING JOB 写入 ──────────────────────────────────

function writeGDCJobs(items) {
  const ss = SpreadsheetApp.openById(SHEET_IDS.gdc);
  let tab = ss.getSheetByName(GDC_JOB_TAB);
  if (!tab) {
    for (const n of ['GDC Job','Marketing Job','Jobs','GDC','Sheet1','工作表1']) {
      tab = ss.getSheetByName(n);
      if (tab) break;
    }
  }
  if (!tab) tab = ss.getSheets()[0];
  if (!tab) return { error: 'No tab found in GDC spreadsheet' };

  const hdrs    = tab.getRange(1, 1, 1, tab.getLastColumn()).getValues()[0];
  const find    = makeFlexFinder(hdrs);
  const numCols = hdrs.length;

  const colMap = {
    id:     find('id'),
    task:   find('任务内容','task','内容','job'),
    date:   find('发布日期','date','日期'),
    dl:     find('截止日期','deadline','dl'),
    cat:    find('类别','cat','type','类型'),
    who:    find('负责人','who','person'),
    pic:    find('pic','截图'),
    status: find('状态','status'),
    pay:    find('pay','费用')
  };

  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2, 1, lastRow - 1, numCols).clearContent();
  if (!items.length) return { ok: true, count: 0 };

  const rows = items.map(item => {
    const row = new Array(numCols).fill('');
    Object.entries(colMap).forEach(([field, idx]) => {
      if (idx >= 0 && item[field] !== undefined) row[idx] = item[field] ?? '';
    });
    return row;
  });

  tab.getRange(2, 1, rows.length, numCols).setValues(rows);
  return { ok: true, count: rows.length };
}

// ─── EVENTS 活动 读取（Task Board spreadsheet）──────────────

function readEventsTab() {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.tasks);
  const tab = ss.getSheetByName(EVENTS_TAB);
  if (!tab) return [];

  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];

  const hdrs = data[0];
  const find = makeFlexFinder(hdrs);

  const col = {
    date:   find('date','日期','活动日期'),
    name:   find('name','活动名称','内容','标题'),
    time:   find('time','时间'),
    type:   find('type','类型'),
    person: find('person','负责人','who'),
    active: find('active','是否启用','启用')
  };

  const g  = (row, c) => c >= 0 ? String(row[c] ?? '').trim() : '';
  const fd = (row, c) => c >= 0 ? fmtDate(row[c]) : '';

  return data.slice(1).filter(row => {
    const d = fd(row, col.date) || g(row, col.date);
    if (!d || !g(row, col.name)) return false;
    if (col.active >= 0) {
      const act = g(row, col.active).toLowerCase();
      if (['false','no','0','✗','x','否'].includes(act)) return false;
    }
    return true;
  }).map((row, idx) => ({
    id:     'ev_s_' + (idx + 1),
    name:   g(row, col.name),
    date:   fd(row, col.date) || g(row, col.date),
    time:   g(row, col.time),
    type:   g(row, col.type) || 'meet',
    person: g(row, col.person)
  }));
}

// ─── ROTATION 轮值 写入 ──────────────────────────────────────

function writeRotationTab(items) {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.tasks);
  let tab   = ss.getSheetByName(ROTATION_TAB);
  if (!tab) {
    tab = ss.insertSheet(ROTATION_TAB);
    tab.getRange(1,1,1,4).setValues([['Date','Activity','Person','Notes']])
       .setFontWeight('bold').setBackground('#fdf8f2');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2, 1, lastRow - 1, 4).clearContent();
  if (!items.length) return { ok: true, count: 0 };
  const rows = items.map(item => [item.date||'', item.act||'', item.person||'', item.theme||'']);
  tab.getRange(2, 1, rows.length, 4).setValues(rows);
  return { ok: true, count: rows.length };
}

// ─── ROTATION 轮值 读取 ──────────────────────────────────────

function readRotationTab() {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.tasks);
  const tab = ss.getSheetByName(ROTATION_TAB);
  if (!tab) return [];

  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];

  const hdrs = data[0];
  const find = makeFlexFinder(hdrs);

  const col = {
    date:   find('date','日期','周四','thursday','week','周次'),
    person: find('person','负责人','who','ylyd person','主持人'),
    note:   find('note','备注')
  };

  const g  = (row, c) => c >= 0 ? String(row[c] ?? '').trim() : '';
  const fd = (row, c) => c >= 0 ? fmtDate(row[c]) : '';

  return data.slice(1).filter(row => {
    return (fd(row, col.date) || g(row, col.date)) && g(row, col.person);
  }).map(row => ({
    date:   fd(row, col.date) || g(row, col.date),
    person: g(row, col.person)
  }));
}

// ─── SALES DASHBOARD ────────────────────────────────────────

function readSalesDashboard() {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.mkt);
  const tab = ss.getSheetByName(SALES_TAB);
  if (!tab) return { error: 'Sales Dashboard tab not found' };

  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];

  let hdrIdx = -1;
  for (let i = 0; i < Math.min(data.length, 15); i++) {
    const first = String(data[i][0] || '').trim();
    if (first === '月份' || first === 'Month') { hdrIdx = i; break; }
  }
  if (hdrIdx < 0) return { error: 'Header row (月份) not found' };

  const hdrs = data[hdrIdx].map(h => String(h).trim());
  const h = {};
  hdrs.forEach((name, i) => { if (name) h[name] = i; });

  const col = (...names) => {
    for (const n of names) { if (h[n] !== undefined) return h[n]; }
    return -1;
  };
  const num = v => parseFloat(String(v||'').replace(/[^0-9.\-]/g,'')) || 0;

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

// ─── 再投资捐款 ──────────────────────────────────────────────

function readDonationTotal() {
  try {
    const ss  = SpreadsheetApp.openById(SHEET_IDS.admin);
    const tab = ss.getSheetByName('再投资捐款');
    if (!tab) return { error: '再投资捐款 tab not found' };

    const data = tab.getDataRange().getValues();
    const num  = v => parseFloat(String(v||'').replace(/[^0-9,\.\-]/g,'').replace(/,/g,'')) || 0;

    for (let i = 0; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        const cell = String(data[i][j]||'').trim();
        if (cell.includes('Total Reinvested') || cell.includes('已捐总额') || cell.includes('Total')) {
          for (let k = data[i].length - 1; k > j; k--) {
            const v = num(data[i][k]);
            if (v > 0) return { total: v };
          }
        }
      }
    }
    return { total: 0 };
  } catch(e) {
    return { error: e.message };
  }
}

// ─── 分钱汇总写入 ────────────────────────────────────────────

const PAY_HEADERS = ['来源','id','任务内容','负责人','类别','类型','截止日期','完成日期','用时','效率','分钱系数','应付金额'];
const COEF_MAP   = {Sales:3,Marketing:2,System:2,GDC:2,'GDC job':2,Operation:1,Admin:1};

function writePaySheet(tasks, mkt, gdc) {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.admin);
  let tab = ss.getSheetByName('分钱汇总');
  if (!tab) {
    tab = ss.insertSheet('分钱汇总');
    tab.getRange(1,1,1,PAY_HEADERS.length).setValues([PAY_HEADERS]).setFontWeight('bold').setBackground('#fdf8f2');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2,1,lastRow-1,PAY_HEADERS.length).clearContent();

  const rows = [];

  (tasks||[]).forEach(t => {
    if (!t.person || t.person==='All' || t.person==='全员') return;
    const eff   = t.eff || calcEff(t.deadline, t.doneDate) || '';
    const coef  = COEF_MAP[t.cat] || 1;
    const hours = parseFloat(t.actual) || 0;
    const pay   = t.type==='fee' ? '' : (hours * coef * (parseFloat(eff)||1)).toFixed(1);
    rows.push(['TaskBoard', t.id||'', t.task||'', t.person||'', t.cat||'', t.type==='fee'?'Fee':'贡献', t.deadline||'', t.doneDate||'', hours||'', eff, coef, pay]);
  });

  (gdc||[]).forEach(t => {
    if (!t.doneDate) return;
    const eff = calcEff(t.dl, t.doneDate) || '';
    rows.push(['GDC', t.id||'', t.task||'', t.who||t.pic||'', 'GDC', 'Fee', t.dl||'', t.doneDate||'', '', eff, 2, '']);
  });

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
    if (type === 'tasks')    return respond(readTaskBoard());
    if (type === 'mkt')      return respond(readMarketingContent());
    if (type === 'gdc')      return respond(readGDCJobs());
    if (type === 'sales')    return respond(readSalesDashboard());
    if (type === 'donation') return respond(readDonationTotal());
    if (type === 'events')   return respond(readEventsTab());
    if (type === 'rotation') return respond(readRotationTab());
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
    if (type === 'tasks')    return respond(writeTaskBoard(Array.isArray(data) ? data : []));
    if (type === 'mkt')      return respond(writeMarketingContent(Array.isArray(data) ? data : []));
    if (type === 'gdc')      return respond(writeGDCJobs(Array.isArray(data) ? data : []));
    if (type === 'rotation') return respond(writeRotationTab(Array.isArray(data) ? data : []));
    if (type === 'pay') {
      return respond(writePaySheet(data.tasks, data.mkt, data.gdc));
    }
    return respond({ error: 'Unknown sheet: ' + type });
  } catch(err) {
    return respond({ error: err.message });
  }
}
