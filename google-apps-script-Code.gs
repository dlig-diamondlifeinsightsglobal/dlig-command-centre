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
const LOGS_TAB          = '自发记录';            // 在 Task Board spreadsheet 里新建这个 tab
const DECISIONS_TAB     = '待决策';              // 在 Task Board spreadsheet 里
const SPEC_MEETINGS_TAB = '特别会议';            // 在 Task Board spreadsheet 里

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
    // 有完成日期 OR 状态列明确写「已完成」，就算已完成
    const rawStatus  = h['状态'] !== undefined ? String(row[h['状态']] ?? '').trim() : '';
    const isRealDone = doneDateRaw !== '' || rawStatus === '已完成';
    const doneDate   = isRealDone ? doneDateRaw : '';
    const status     = isRealDone ? '已完成' : (rawStatus || '待启动');

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
    if (h['状态']     !== undefined) row[h['状态']]     = item.status   || '';
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
    id:       find('id'),
    task:     find('发布内容','任务内容','task','内容','job','工作'),
    date:     find('发布日期','date','日期'),
    dl:       find('截止日期','deadline','dl','due'),
    platform: find('平台','platform'),
    cat:      find('类别','cat','type','类型','category'),
    who:      find('gdc对接人','gdc对接','gdc 对接','对接人','负责 gdc','who'),
    pic:      find('负责人','dlig跟进','dlig负责人','dlig负责','dlig 跟进','跟进','pic','截图','person'),
    doneDate: find('完成日期','donedate','done date','done','完成'),
    eff:      find('效率','eff','efficiency'),
    status:   find('状态','status'),
    pay:      find('pay','费用','amount','金额')
  };

  const g  = (row, c) => c >= 0 ? String(row[c] ?? '').trim() : '';
  const fd = (row, c) => c >= 0 ? fmtDate(row[c]) : '';

  return data.slice(1)
    .filter(r => col.task >= 0 && String(r[col.task] || '').trim() !== '')
    .map((row, idx) => {
      const doneDate = fd(row, col.doneDate) || g(row, col.doneDate);
      const statusRaw = g(row, col.status);
      // derive status from doneDate if no status column
      const status = statusRaw || (doneDate ? '已完成' : '待开始');
      return {
        id:       g(row, col.id) || ('gdc_' + g(row, col.task).slice(0,30).replace(/\s+/g,'_')),
        task:     g(row, col.task),
        date:     fd(row, col.date)   || g(row, col.date),
        dl:       fd(row, col.dl)     || g(row, col.dl),
        platform: g(row, col.platform),
        cat:      g(row, col.cat),
        who:      g(row, col.who),
        pic:      g(row, col.pic),
        doneDate: doneDate,
        eff:      g(row, col.eff),
        status:   status,
        pay:      g(row, col.pay)
      };
    });
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
    id:       find('id'),
    task:     find('发布内容','任务内容','task','内容','job'),
    date:     find('发布日期','date','日期'),
    dl:       find('截止日期','deadline','dl'),
    platform: find('平台','platform'),
    cat:      find('类别','cat','type','类型'),
    who:      find('gdc对接人','gdc对接','gdc 对接','对接人','负责 gdc','who'),
    pic:      find('负责人','dlig跟进','dlig负责人','dlig负责','dlig 跟进','跟进','pic','截图','person'),
    doneDate: find('完成日期','donedate','done date','done','完成'),
    eff:      find('效率','eff','efficiency'),
    status:   find('状态','status'),
    pay:      find('pay','费用')
  };

  // Only write up to the highest mapped column — avoids wiping formula/button columns (e.g. 操作)
  const mappedIdxs = Object.values(colMap).filter(v => v >= 0);
  const writeCols = mappedIdxs.length > 0 ? Math.max(...mappedIdxs) + 1 : numCols;

  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2, 1, lastRow - 1, writeCols).clearContent();
  if (!items.length) return { ok: true, count: 0 };

  const rows = items.map(item => {
    const row = new Array(writeCols).fill('');
    Object.entries(colMap).forEach(([field, idx]) => {
      if (idx >= 0 && idx < writeCols && item[field] !== undefined) row[idx] = item[field] ?? '';
    });
    return row;
  });

  tab.getRange(2, 1, rows.length, writeCols).setValues(rows);
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
    repeat: find('repeat','重复','循环'),
    active: find('active','是否启用','启用'),
    mode:   find('mode','形式','online/physical'),
    loc:    find('loc','location','地点','链接')
  };

  const g  = (row, c) => c >= 0 ? String(row[c] ?? '').trim() : '';
  const fd = (row, c) => c >= 0 ? fmtDate(row[c]) : '';

  // Send ALL rows to frontend (including active=false — used as cancellation markers for repeat events)
  return data.slice(1).filter(row => {
    const d = fd(row, col.date) || g(row, col.date);
    return d && g(row, col.name); // only skip completely empty rows
  }).map((row, idx) => {
    const actRaw = g(row, col.active).toLowerCase();
    const active = !['false','no','0','✗','x','否'].includes(actRaw);
    return {
      id:     'ev_s_' + (idx + 1),
      name:   g(row, col.name),
      date:   fd(row, col.date) || g(row, col.date),
      time:   g(row, col.time),
      type:   g(row, col.type) || 'meet',
      person: g(row, col.person),
      repeat: g(row, col.repeat).toLowerCase(),
      active: active,
      mode:   g(row, col.mode).toLowerCase(),
      loc:    g(row, col.loc)
    };
  });
}

// ─── EVENTS 活动 写入 ───────────────────────────────────────

function writeEventsTab(items) {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.tasks);
  let tab   = ss.getSheetByName(EVENTS_TAB);
  if (!tab) {
    tab = ss.insertSheet(EVENTS_TAB);
  }
  const HDR = ['date','name','time','type','person','repeat','active','mode','loc'];
  tab.clearContents();
  tab.getRange(1, 1, 1, HDR.length).setValues([HDR])
     .setFontWeight('bold').setBackground('#e8f4fd');
  if (!Array.isArray(items) || items.length === 0) return { ok: true, rows: 0 };
  const rows = items.map(e => [
    e.date   || '',
    e.name   || '',
    e.time   || '',
    e.type   || 'meet',
    e.person || '',
    e.repeat || '',
    e.active === false ? 'FALSE' : 'TRUE',
    e.mode   || '',
    e.loc    || ''
  ]);
  tab.getRange(2, 1, rows.length, HDR.length).setValues(rows);
  return { ok: true, rows: rows.length };
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

// ─── 自发记录 ──────────────────────────────────────────────────

function readLogs() {
  const ss = SpreadsheetApp.openById(SHEET_IDS.tasks);
  let sh = ss.getSheetByName(LOGS_TAB);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const hdrs = rows[0].map(h => String(h).trim().toLowerCase());
  const find = makeFlexFinder(rows[0]);
  const iId = find('id'); const iDesc = find('desc','描述','做了什么');
  const iPerson = find('person','负责人'); const iCat = find('cat','类别','category');
  const iHours = find('hours','实际时间','时间'); const iContrib = find('contrib','贡献值');
  const iTs = find('ts','timestamp','时间戳');
  return rows.slice(1).filter(r => r[iDesc]).map(r => ({
    id:      String(r[iId]  || ''),
    desc:    String(r[iDesc]|| ''),
    person:  String(r[iPerson]>=0 ? r[iPerson] : ''),
    cat:     String(r[iCat] >=0 ? r[iCat]  : ''),
    hours:   parseFloat(r[iHours]) || 0,
    contrib: parseFloat(r[iContrib])|| 0,
    ts:      r[iTs] ? Number(r[iTs]) : 0
  }));
}

function writeLogs(rows) {
  const ss = SpreadsheetApp.openById(SHEET_IDS.tasks);
  let sh = ss.getSheetByName(LOGS_TAB);
  if (!sh) { sh = ss.insertSheet(LOGS_TAB); }
  sh.clearContents();
  const hdrs = ['id','desc','person','cat','hours','contrib','ts'];
  sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
  if (rows.length > 0) {
    const data = rows.map(r => [r.id||'', r.desc||'', r.person||'', r.cat||'', r.hours||0, r.contrib||0, r.ts||0]);
    sh.getRange(2, 1, data.length, hdrs.length).setValues(data);
  }
  return { ok: true, count: rows.length };
}

// ─── 待决策 读写 ──────────────────────────────────────────────

function readDecisions() {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.tasks);
  const tab = ss.getSheetByName(DECISIONS_TAB);
  if (!tab) return [];
  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(r => r[0]).map(r => {
    let opts=[], votes={}, comments=[];
    try { opts     = JSON.parse(String(r[2]||'[]')); } catch(e){}
    try { votes    = JSON.parse(String(r[3]||'{}')); } catch(e){}
    try { comments = JSON.parse(String(r[4]||'[]')); } catch(e){}
    const type     = String(r[7]||'vote');
    const resolved = String(r[8]||'') === 'true';
    return { id:String(r[0]), q:String(r[1]||''), opts, votes, comments,
             owner:String(r[5]||''), dl:String(r[6]||'待定'), type, resolved };
  });
}

function writeDecisions(items) {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.tasks);
  let tab   = ss.getSheetByName(DECISIONS_TAB);
  if (!tab) {
    tab = ss.insertSheet(DECISIONS_TAB);
    tab.getRange(1,1,1,9).setValues([['id','question','options','votes','comments','owner','deadline','type','resolved']])
       .setFontWeight('bold').setBackground('#fff0f5');
  } else {
    // 确保 header 有 type/resolved 列
    const h = tab.getRange(1,1,1,9).getValues()[0];
    if (!h[7]) tab.getRange(1,8).setValue('type');
    if (!h[8]) tab.getRange(1,9).setValue('resolved');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2,1,lastRow-1,9).clearContent();
  if (!items.length) return { ok:true, count:0 };
  const rows = items.map(d => [
    d.id||'', d.q||'',
    JSON.stringify(d.opts||[]),
    JSON.stringify(d.votes||{}),
    JSON.stringify(d.comments||[]),
    d.owner||'', d.dl||'待定',
    d.type||'vote',
    d.resolved ? 'true' : 'false'
  ]);
  tab.getRange(2,1,rows.length,9).setValues(rows);
  return { ok:true, count:rows.length };
}


// ─── 特别会议 读写 ────────────────────────────────────────────

function readSpecialMeetings() {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.tasks);
  const tab = ss.getSheetByName(SPEC_MEETINGS_TAB);
  if (!tab) return [];
  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(r => r[0]).map(r => {
    let agenda = [];
    try { agenda = JSON.parse(String(r[5]||'[]')); } catch(e){}
    return { id:String(r[0]), title:String(r[1]||''),
             date:fmtDate(r[2])||String(r[2]||''), time:String(r[3]||''),
             loc:String(r[4]||''), agenda };
  });
}

function writeSpecialMeetings(items) {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.tasks);
  let tab   = ss.getSheetByName(SPEC_MEETINGS_TAB);
  if (!tab) {
    tab = ss.insertSheet(SPEC_MEETINGS_TAB);
    tab.getRange(1,1,1,6).setValues([['id','title','date','time','location','agenda']])
       .setFontWeight('bold').setBackground('#fff5f5');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2,1,lastRow-1,6).clearContent();
  if (!items.length) return { ok:true, count:0 };
  const rows = items.map(m => [
    m.id||'', m.title||'', m.date||'', m.time||'', m.loc||'',
    JSON.stringify(m.agenda||[])
  ]);
  tab.getRange(2,1,rows.length,6).setValues(rows);
  return { ok:true, count:rows.length };
}

// ─── 心动觉察 出席者名单 ──────────────────────────────────────

function readAttendance() {
  const ss  = getOpsSS();
  const tab = ss.getSheetByName('出席者名单');
  if (!tab) return {};
  const data = tab.getDataRange().getValues();
  if (data.length < 2) return {};
  const result = {};
  data.slice(1).filter(r => r[0]).forEach(r => {
    const session = String(r[0]);
    if (!result[session]) result[session] = [];
    result[session].push({
      name:  String(r[1]||''),
      phone: String(r[2]||''),
      rel:   String(r[3]||''),
      zone:  String(r[4]||''),
      d1:    r[5]===true||r[5]==='TRUE'||r[5]==='true',
      d2:    r[6]===true||r[6]==='TRUE'||r[6]==='true'
    });
  });
  return result;
}

function writeAttendance(obj) {
  const ss  = getOpsSS();
  let tab   = ss.getSheetByName('出席者名单');
  if (!tab) {
    tab = ss.insertSheet('出席者名单');
    tab.getRange(1,1,1,7).setValues([['场次','姓名','电话','关系状态','迎宾分组','Day1','Day2']])
       .setFontWeight('bold').setBackground('#f0fdf4');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2,1,lastRow-1,7).clearContent();
  const rows = [];
  Object.entries(obj||{}).forEach(([session, list]) => {
    (list||[]).forEach(a => rows.push([session, a.name||'', a.phone||'', a.rel||'', a.zone||'', a.d1?'TRUE':'FALSE', a.d2?'TRUE':'FALSE']));
  });
  if (rows.length > 0) tab.getRange(2,1,rows.length,7).setValues(rows);
  return { ok:true, count:rows.length };
}

// ─── 试炼记录：沙盘 & 拍卖 ────────────────────────────────────

function readSbRecords() {
  const ss  = getOpsSS();
  const tab = ss.getSheetByName('沙盘试炼记录');
  if (!tab) return [];
  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(r => r[0]).map(r => ({
    unit:  String(r[0]||''), p1:String(r[1]||''), p2:String(r[2]||''),
    dream: String(r[3]||''), goal:String(r[4]||''), note:String(r[5]||'')
  }));
}

function writeSbRecords(items) {
  const ss  = getOpsSS();
  let tab   = ss.getSheetByName('沙盘试炼记录');
  if (!tab) {
    tab = ss.insertSheet('沙盘试炼记录');
    tab.getRange(1,1,1,6).setValues([['Unit','Player1','Player2','梦想目标','人生目标','觉察备注']])
       .setFontWeight('bold').setBackground('#fffbeb');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2,1,lastRow-1,6).clearContent();
  if (!items.length) return { ok:true, count:0 };
  const rows = items.map(r => [r.unit||'', r.p1||'', r.p2||'', r.dream||'', r.goal||'', r.note||'']);
  tab.getRange(2,1,rows.length,6).setValues(rows);
  return { ok:true, count:rows.length };
}

function readAuRecords() {
  const ss  = getOpsSS();
  const tab = ss.getSheetByName('拍卖试炼记录');
  if (!tab) return [];
  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(r => r[0]).map(r => ({
    trait:  String(r[0]||''), price:String(r[1]||''), buyer:String(r[2]||''),
    seller: String(r[3]||''), benef:String(r[4]||''), note:String(r[5]||'')
  }));
}

function writeAuRecords(items) {
  const ss  = getOpsSS();
  let tab   = ss.getSheetByName('拍卖试炼记录');
  if (!tab) {
    tab = ss.insertSheet('拍卖试炼记录');
    tab.getRange(1,1,1,6).setValues([['特质','成交价','拍到者','给予者','受益者','备注']])
       .setFontWeight('bold').setBackground('#fffbeb');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2,1,lastRow-1,6).clearContent();
  if (!items.length) return { ok:true, count:0 };
  const rows = items.map(r => [r.trait||'', r.price||'', r.buyer||'', r.seller||'', r.benef||'', r.note||'']);
  tab.getRange(2,1,rows.length,6).setValues(rows);
  return { ok:true, count:rows.length };
}

// ─── 岗位负责人 读写 ──────────────────────────────────────────

function readXdRoles() {
  const ss  = getOpsSS();
  const tab = ss.getSheetByName('岗位负责人');
  if (!tab) return null;
  const data = tab.getDataRange().getValues();
  if (data.length < 2) return null;
  const d1 = [], d2 = [];
  data.slice(1).filter(r => r[0]).forEach(r => {
    const day = String(r[0]);
    const obj = { role:String(r[1]||''), detail:String(r[2]||''), person:String(r[3]||'') };
    if (day === 'D1') d1.push(obj); else d2.push(obj);
  });
  return { d1, d2 };
}

function writeXdRoles(obj) {
  const ss  = getOpsSS();
  let tab   = ss.getSheetByName('岗位负责人');
  if (!tab) {
    tab = ss.insertSheet('岗位负责人');
    tab.getRange(1,1,1,4).setValues([['Day','岗位','职责','负责人']])
       .setFontWeight('bold').setBackground('#eff6ff');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2,1,lastRow-1,4).clearContent();
  const rows = [];
  (obj.d1||[]).forEach(r => rows.push(['D1', r.role||'', r.detail||'', r.person||'']));
  (obj.d2||[]).forEach(r => rows.push(['D2', r.role||'', r.detail||'', r.person||'']));
  if (rows.length > 0) tab.getRange(2,1,rows.length,4).setValues(rows);
  return { ok:true, count:rows.length };
}

// ─── INVENTORY 读取（Admin spreadsheet Inventory tab）──────────

function readInventory() {
  const ss  = SpreadsheetApp.openById(SHEET_IDS.admin);
  let tab   = ss.getSheetByName('Inventory');
  if (!tab) tab = ss.getSheetByName('inventory');
  if (!tab) tab = ss.getSheetByName('库存');
  if (!tab) return { error: 'Inventory tab not found' };

  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];

  const hdrs = data[0];
  const find = makeFlexFinder(hdrs);

  const col = {
    n:    find('物品名称','名称','item','name'),
    cat:  find('分类','category','类别','cat'),
    qty:  find('库存量','库存','qty','quantity','数量'),
    min:  find('最低量','最低','minimum','min'),
    s:    find('状态','status','state'),
    lead: find('到货 lead time','lead time','leadtime','到货','工作日'),
    link: find('购买链接','link','url','购买','链接')
  };

  const g = (row, c) => c >= 0 ? String(row[c] ?? '').trim() : '';

  return data.slice(1)
    .filter(r => col.n >= 0 && String(r[col.n] || '').trim() !== '')
    .map(row => ({
      n:    g(row, col.n),
      cat:  g(row, col.cat),
      qty:  g(row, col.qty) !== '' ? Number(g(row, col.qty)) || 0 : null,
      min:  g(row, col.min) !== '' ? Number(g(row, col.min)) || 0 : 0,
      s:    g(row, col.s)   || '充足',
      lead: g(row, col.lead),
      link: g(row, col.link)
    }));
}

// ─── 书影充电站 (Operation Sheet) ────────────────────────────
const OPS_SHEET_ID  = '1_6bghVk44SlC9tH9vDENcw3JeorzCINaWSbkMbMcS38';
const BS_FLOW_TAB   = '书影充电站流程';
const BS_ROT_TAB    = '👤 轮值';

function getOpsSS() { return SpreadsheetApp.openById(OPS_SHEET_ID); }

// Combined tab: top section = settings (key|value), blank row, then flow table header + rows
function readBSCombined() {
  const ss = getOpsSS();
  let tab = ss.getSheetByName(BS_FLOW_TAB);
  if (!tab) {
    tab = ss.insertSheet(BS_FLOW_TAB);
    const allRows = [
      ['zoom_link',   'https://us06web.zoom.us/j/7707709233', '', ''],
      ['zoom_id',     '7707709233', '', ''],
      ['zoom_passcode','DLIG', '', ''],
      ['zoom_host_key','779233', '', ''],
      ['ppt_link',    'https://canva.link/ylydwisdomrechargesat', '', ''],
      ['wisdomlib_link','https://diamondlifeinsightsglobal.app.clientclub.net/courses/products/4a1a5ae9-69c5-4c22-825f-b2dd26402a1a?source=courses', '', ''],
      ['other_info',  '', '', ''],
      ['', '', '', ''],
      ['时间','时长','环节','说明'],
      ['2:30 PM','15min','开场 & 充电模式','开场 & 自我介绍在 chatbox'],
      ['2:45 PM','45min','📚 书影分享 / Lucky draw','video / 《超凡的智慧4问题》 / 《爱种子问题》· 模式1—Q&A 或 模式2—见证分享（5min分享 + 10min拆解：总结2个 action keypoint，分享可变成广告）'],
      ['3:30 PM','10min','Q&A 或见证分享、合照','复盘 & 拆解 keypoints'],
      ['3:40 PM','15min','Me Time 复盘','YLYD 设计册 · 偶尔提 Why Me Time ❤'],
      ['3:55 PM','5min','集体乐咖','好种子库 ❤'],
      ['4:00 PM','—','END','带着智慧，回归生活']
    ];
    tab.getRange(1,1,allRows.length,4).setValues(allRows);
    tab.getRange(1,1,7,1).setFontWeight('bold').setFontColor('#7c3aed');
    tab.getRange(9,1,1,4).setFontWeight('bold').setBackground('#e8f4fd');
    tab.setColumnWidth(1,150); tab.setColumnWidth(2,60); tab.setColumnWidth(3,180); tab.setColumnWidth(4,400);
    return {
      settings:{zoom_link:'https://us06web.zoom.us/j/7707709233',zoom_id:'7707709233',zoom_passcode:'DLIG',zoom_host_key:'779233',ppt_link:'https://canva.link/ylydwisdomrechargesat',wisdomlib_link:'https://diamondlifeinsightsglobal.app.clientclub.net/courses/products/4a1a5ae9-69c5-4c22-825f-b2dd26402a1a?source=courses',other_info:''},
      flow:[{time:'2:30 PM',duration:'15min',section:'开场 & 充电模式',desc:'开场 & 自我介绍在 chatbox'},{time:'2:45 PM',duration:'45min',section:'📚 书影分享 / Lucky draw',desc:'video / 《超凡的智慧4问题》'},{time:'3:30 PM',duration:'10min',section:'Q&A 或见证分享、合照',desc:'复盘 & 拆解 keypoints'},{time:'3:40 PM',duration:'15min',section:'Me Time 复盘',desc:'YLYD 设计册'},{time:'3:55 PM',duration:'5min',section:'集体乐咖',desc:'好种子库 ❤'},{time:'4:00 PM',duration:'—',section:'END',desc:'带着智慧，回归生活'}]
    };
  }
  const data = tab.getDataRange().getValues();
  const settingsKeys = ['zoom_link','zoom_id','zoom_passcode','zoom_host_key','ppt_link','wisdomlib_link','other_info'];
  const settings = {};
  let flowStartRow = -1;
  const tz = Session.getScriptTimeZone();
  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0]||'').trim();
    if (settingsKeys.includes(key)) { settings[key] = String(data[i][1]||'').trim(); }
    if (key === '时间') { flowStartRow = i + 1; }
  }
  const flow = flowStartRow >= 0 ? data.slice(flowStartRow).filter(r=>r[0]).map(r=>({
    time: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'h:mm a') : String(r[0]||''),
    duration: String(r[1]||''), section: String(r[2]||''), desc: String(r[3]||'')
  })) : [];
  return { settings, flow };
}

function readBSRotation() {
  const ss = getOpsSS();
  const tab = ss.getSheetByName(BS_ROT_TAB);
  if (!tab) return [];
  const data = tab.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(r=>r[0]).map(r=>({
    date: fmtDate(r[0])||String(r[0]||''), person:String(r[1]||''), notes:String(r[2]||'')
  }));
}

function writeBSRotation(items) {
  const ss = getOpsSS();
  let tab = ss.getSheetByName(BS_ROT_TAB);
  if (!tab) {
    tab = ss.insertSheet(BS_ROT_TAB);
    tab.getRange(1,1,1,3).setValues([['date','person','notes']]).setFontWeight('bold').setBackground('#f0fdf4');
  }
  const lastRow = tab.getLastRow();
  if (lastRow > 1) tab.getRange(2,1,lastRow-1,3).clearContent();
  if (!items.length) return {ok:true,count:0};
  tab.getRange(2,1,items.length,3).setValues(items.map(r=>[r.date||'',r.person||'',r.notes||'']));
  return {ok:true,count:items.length};
}

function readBSAll() {
  const combined = readBSCombined();
  return {settings:combined.settings, flow:combined.flow, rotation:readBSRotation()};
}

// ═══════════════════════════════════════════════════════════════
// DLIG Task Tracker — 分钱计算工具
// 在 Script Editor 运行一次 buildDLIGTracker()，自动在 Admin Sheet 生成所有 tabs
// Admin Sheet: https://docs.google.com/spreadsheets/d/1zGy3rV0bv2dERFRWoRFhRLGj68oJ3RtAj-kiqi71DXk
// ═══════════════════════════════════════════════════════════════

function buildDLIGTracker() {
  const ss = SpreadsheetApp.openById(SHEET_IDS.tasks);

  // 安全地获取或新建 sheet（不删除其他已有 tabs）
  function getOrCreate(name) {
    return ss.getSheetByName(name) || ss.insertSheet(name);
  }

  const lkpSheet = getOrCreate("⚙️ Lookup");
  const logSheet = getOrCreate("📝 Task Log");
  const paySheet = getOrCreate("💰 月结算 Payout");
  const refSheet = getOrCreate("📋 参考 Reference");

  buildLookup(lkpSheet);
  buildTaskLog(ss, logSheet, lkpSheet);
  buildPayout(ss, paySheet, logSheet);
  buildReference(refSheet);

  // 隐藏 Lookup（Google Sheets 可从隐藏 sheet 读取 data validation）
  lkpSheet.hideSheet();

  // 跳到 Task Log
  ss.setActiveSheet(logSheet);
  SpreadsheetApp.getUi().alert("✅ DLIG Task Tracker 已生成！\n\n平时用 📝 Task Log 记录\n月底用 💰 月结算 Payout 算钱");
}

// ─── COLOURS ────────────────────────────────────────────────────
const C = {
  dblue:  "#1F4E79", lblue:  "#D6E4F0", white:  "#FFFFFF",
  black:  "#000000", dgold:  "#B8860B", lgray:  "#F5F5F5",
  mgray:  "#D9D9D9", input:  "#CCE5FF", yellow: "#FFFF99",
  green:  "#E2EFDA", dgreen: "#375623", purple: "#EAE0F0",
  dpurp:  "#5C2D8E", orange: "#FCE4D6", doran:  "#833C00",
  teal:   "#1F6B75", lteal:  "#D9F0F2", red:    "#FFC7CE",
};

function bg(sheet, r, c, color)  { sheet.getRange(r,c).setBackground(color); }
function bgR(sheet, r, c1, c2, color) { sheet.getRange(r,c1,1,c2-c1+1).setBackground(color); }

function hdrCell(sheet, r, c, text, bgColor, fgColor="#FFFFFF", bold=true, size=10, hMerge=1, wrap=false) {
  const range = sheet.getRange(r, c, 1, hMerge);
  if (hMerge > 1) range.merge();
  range.setValue(text)
    .setBackground(bgColor)
    .setFontColor(fgColor)
    .setFontWeight(bold ? "bold" : "normal")
    .setFontSize(size)
    .setFontFamily("Arial")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(wrap);
  return range;
}

function lblCell(sheet, r, c, text, bgColor, bold=false, fgColor="#000000", size=10, align="left", hMerge=1, wrap=false) {
  const range = sheet.getRange(r, c, 1, hMerge);
  if (hMerge > 1) range.merge();
  range.setValue(text)
    .setBackground(bgColor)
    .setFontColor(fgColor)
    .setFontWeight(bold ? "bold" : "normal")
    .setFontSize(size)
    .setFontFamily("Arial")
    .setHorizontalAlignment(align)
    .setVerticalAlignment("middle")
    .setWrap(wrap);
  return range;
}

function frmCell(sheet, r, c, formula, fmt, bgColor, bold=false, fgColor="#000000") {
  const cell = sheet.getRange(r, c);
  cell.setFormula(formula)
    .setBackground(bgColor)
    .setFontColor(fgColor)
    .setFontWeight(bold ? "bold" : "normal")
    .setFontFamily("Arial")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  if (fmt) cell.setNumberFormat(fmt);
  return cell;
}

function inpCell(sheet, r, c, value, fmt, bgColor=C.input) {
  const cell = sheet.getRange(r, c);
  if (value !== null && value !== undefined) cell.setValue(value);
  cell.setBackground(bgColor)
    .setFontColor("#00008B")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  if (fmt) cell.setNumberFormat(fmt);
  return cell;
}

const RM  = '"RM "#,##0.00';
const PCT = '0.0%';

// ─── LOOKUP SHEET ───────────────────────────────────────────────
function buildLookup(sh) {
  sh.clearContents();

  // DLIG Tasks: col A=name, B=price, C=category_key
  const DLIG_TASKS = [
    ["视频文案（2篇）",           5,   "DLIG 视频文案"],
    ["信息文案模版（2篇）",        5,   "DLIG 文案模版"],
    ["制作Post图+文案（1套）",    15,   "DLIG Post图文"],
    ["制作海报（1张）",           15,   "DLIG 海报"],
    ["视频脚本（1支）",            5,   "DLIG 视频脚本"],
    ["剪辑短视频-一键成片",        20,   "DLIG 一键成片"],
    ["剪辑短视频（≤1min）",       50,   "DLIG 短视频剪"],
    ["剪辑YouTube（≤25min）",    100,   "DLIG YouTube剪"],
    ["Landing Page文案",         50,   "DLIG LP文案"],
    ["Landing Page制作",        150,   "DLIG LP制作"],
    ["线下活动带领（半天）",        50,   "DLIG 活动带领"],
    ["副村长-带盘（半天）",         30,   "DLIG 副村长带盘"],
    ["Crew-无带盘（半天）",         20,   "DLIG Crew"],
  ];

  // GDC Tasks: col E=name, F=coefficient
  const GDC_TASKS = [
    ["GDC 文案书写",   2],
    ["GDC 视频文案",   2],
    ["GDC 图文Post",   3],
    ["GDC 短视频脚本", 1],
    ["GDC 短视频剪",   4],
    ["GDC job",        0],
  ];

  const MEMBERS = ["Li Joo", "Stella", "Roy", "Jasper"];

  // Write headers
  sh.getRange(1,1).setValue("DLIG_Task");
  sh.getRange(1,2).setValue("单价");
  sh.getRange(1,3).setValue("类别Key");
  sh.getRange(1,5).setValue("GDC_Cat");
  sh.getRange(1,6).setValue("系数");
  sh.getRange(1,8).setValue("Members");
  sh.getRange(1,10).setValue("AllCats");

  // Write DLIG tasks
  DLIG_TASKS.forEach((row, i) => {
    sh.getRange(i+2, 1).setValue(row[0]);
    sh.getRange(i+2, 2).setValue(row[1]);
    sh.getRange(i+2, 3).setValue(row[2]);
  });

  // Write GDC tasks
  GDC_TASKS.forEach((row, i) => {
    sh.getRange(i+2, 5).setValue(row[0]);
    sh.getRange(i+2, 6).setValue(row[1]);
  });

  // Write members
  MEMBERS.forEach((m, i) => sh.getRange(i+2, 8).setValue(m));

  // AllCats = DLIG category keys + GDC names + 其他
  const allCats = DLIG_TASKS.map(t => t[2])
    .concat(GDC_TASKS.map(t => t[0]))
    .concat(["其他（记时间）"]);
  allCats.forEach((c, i) => sh.getRange(i+2, 10).setValue(c));

  // Store counts in named cells for reference
  sh.getRange(1,12).setValue("DLIG_COUNT"); sh.getRange(1,13).setValue(DLIG_TASKS.length);
  sh.getRange(2,12).setValue("GDC_COUNT");  sh.getRange(2,13).setValue(GDC_TASKS.length);
  sh.getRange(3,12).setValue("MEM_COUNT");  sh.getRange(3,13).setValue(MEMBERS.length);
  sh.getRange(4,12).setValue("CAT_COUNT");  sh.getRange(4,13).setValue(allCats.length);
}

// ─── TASK LOG ───────────────────────────────────────────────────
function buildTaskLog(ss, sh, lkpSheet) {
  sh.clearContents();
  sh.clearFormats();

  const LOG_ROWS = 300;
  const lkpName  = lkpSheet.getName();

  // Column widths: A=5 B=13 C=35 D=14 E=22 F=9 G=11 H=11 I=15 J=15 K=20
  [50,90,260,110,180,70,90,90,120,120,160].forEach((w,i) => sh.setColumnWidth(i+1, w));
  sh.setFrozenRows(4);

  let r = 1;

  // Title
  hdrCell(sh, r, 1, "💎 DLIG — Task 记录表", C.dblue, C.white, true, 13, 11);
  sh.setRowHeight(r, 36); r++;

  // Subtitle
  hdrCell(sh, r, 1, "✏️ 负责人和类别请用下拉选择。单价和小计自动算。效率系数：准时=1 / 提早=1.5 / 延迟=0.5 / 未完成=0", C.dblue, C.white, false, 9, 11);
  sh.setRowHeight(r, 16); r++;

  // Legend
  sh.getRange(r,1).setBackground(C.input).setValue(" ").setHorizontalAlignment("center");
  sh.getRange(r,2).setValue("蓝=填写").setFontSize(9).setFontStyle("italic");
  sh.getRange(r,3).setBackground(C.lgray).setValue(" ").setHorizontalAlignment("center");
  sh.getRange(r,4).setValue("灰=自动算").setFontSize(9).setFontStyle("italic");
  sh.getRange(r,5).setBackground(C.yellow).setValue(" ").setHorizontalAlignment("center");
  sh.getRange(r,6).setValue("黄=GDC系数").setFontSize(9).setFontStyle("italic");
  sh.getRange(r,7,1,5).merge().setValue("❗ GDC task: 系数自动带出 | DLIG task: 单价自动带出").setFontSize(9).setFontColor(C.doran);
  sh.setRowHeight(r, 14); r++;

  // Headers
  const hdrs = ["#","日期","任务内容","负责人","类别","用时(hrs)","效率系数","GDC系数","单价(RM)","小计(RM)","备注"];
  hdrs.forEach((h, i) => hdrCell(sh, r, i+1, h, C.dgold, C.white, true, 10, 1, true));
  sh.setRowHeight(r, 26);
  const HDR_ROW = r;
  const DATA_START = r + 1;
  r++;

  // Data validation ranges in Lookup
  const memberRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(lkpName).getRange("H2:H5"), true)
    .setAllowInvalid(false).build();
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(lkpName).getRange("J2:J23"), true)
    .setAllowInvalid(false).build();

  for (let i = 0; i < LOG_ROWS; i++) {
    const bg0 = i % 2 === 0 ? C.lgray : C.white;
    const bg1 = i % 2 === 0 ? C.input : "#E8F4FF";

    // A: row number
    frmCell(sh, r, 1, `=IF(C${r}="","",ROW()-${HDR_ROW})`, null, bg0);
    // B: date
    inpCell(sh, r, 2, null, "DD/MM/YYYY", bg1);
    // C: task description
    sh.getRange(r,3).setBackground(bg1).setFontColor("#00008B").setHorizontalAlignment("left").setVerticalAlignment("middle");
    // D: member dropdown
    sh.getRange(r,4).setBackground(bg1).setDataValidation(memberRule).setFontColor("#00008B").setHorizontalAlignment("center").setVerticalAlignment("middle");
    // E: category dropdown
    sh.getRange(r,5).setBackground(bg1).setDataValidation(catRule).setFontColor("#00008B").setHorizontalAlignment("center").setVerticalAlignment("middle");
    // F: hours
    inpCell(sh, r, 6, null, "0.0", bg1);
    // G: efficiency coefficient
    inpCell(sh, r, 7, 1, "0.0", bg1);
    // H: GDC coefficient (auto)
    frmCell(sh, r, 8,
      `=IF(E${r}="","",IFERROR(VLOOKUP(E${r},'${lkpName}'!$E:$F,2,0),""))`,
      "0.0", C.yellow, false, "#7B3F00");
    // I: unit price (auto, DLIG only)
    frmCell(sh, r, 9,
      `=IF(E${r}="","",IFERROR(INDEX('${lkpName}'!$B:$B,MATCH(E${r},'${lkpName}'!$C:$C,0)),""))`,
      RM, bg0);
    // J: subtotal (DLIG: price × eff | GDC: blank — calculated in payout)
    frmCell(sh, r, 10,
      `=IF(I${r}="","",I${r}*G${r})`,
      RM, bg0, true);
    // K: notes
    sh.getRange(r,11).setBackground(bg1).setFontColor("#00008B").setVerticalAlignment("middle");

    sh.setRowHeight(r, 20);
    r++;
  }

  const DATA_END = r - 1;

  // Store data range info in a named range for payout sheet
  ss.setNamedRange("LOG_DATA_START", sh.getRange(DATA_START, 1));
  ss.setNamedRange("LOG_DATA_END",   sh.getRange(DATA_END, 1));
}

// ─── PAYOUT SHEET ───────────────────────────────────────────────
function buildPayout(ss, sh, logSheet) {
  sh.clearContents();
  sh.clearFormats();

  const logName = logSheet.getName();
  // DATA_START = row 5 (1 title + 1 subtitle + 1 legend + 1 header + 1 = row 5)
  const DS = 5;
  const DE = 304; // DS + LOG_ROWS - 1

  [20,180,140,140,140,140,140].forEach((w,i) => sh.setColumnWidth(i+1, w));

  const MEMBERS = ["Li Joo","Stella","Roy","Jasper"];
  let r = 1;

  hdrCell(sh, r, 2, "💰 DLIG — 月结算 Payout", C.dblue, C.white, true, 13, 6);
  sh.setRowHeight(r, 32); r++;
  hdrCell(sh, r, 2, "只需填写月份、Sales Target、GDC Project收入，其余全部自动计算。", C.dblue, C.white, false, 9, 6);
  sh.setRowHeight(r, 16); r+=2;

  // ── A. 基本设置
  hdrCell(sh, r, 2, "⚙️ 基本设置", C.dblue, C.white, true, 11, 6); sh.setRowHeight(r, 24); r++;

  lblCell(sh, r, 2, "结算月份", C.lblue, true);
  inpCell(sh, r, 3, "2026-05");
  lblCell(sh, r, 4, "格式 YYYY-MM", C.lgray, false, C.black, 9, "left", 4);
  sh.setRowHeight(r, 22); const MONTH_ROW = r; r++;

  lblCell(sh, r, 2, "Sales Target (RM)", C.lblue, true);
  inpCell(sh, r, 3, 0, RM);
  lblCell(sh, r, 4, "本月目标业绩", C.lgray, false, C.black, 9, "left", 4);
  sh.setRowHeight(r, 22); const TARGET_ROW = r; r++;

  lblCell(sh, r, 2, "实际销售额 (RM)", C.lblue, true);
  inpCell(sh, r, 3, 0, RM);
  lblCell(sh, r, 4, "本月实际达成", C.lgray, false, C.black, 9, "left", 4);
  sh.setRowHeight(r, 22); const ACTUAL_ROW = r; r++;

  lblCell(sh, r, 2, "DLIG Task 达成率", C.lblue, true);
  frmCell(sh, r, 3, `=IF(C${TARGET_ROW}=0,0,MIN(C${ACTUAL_ROW}/C${TARGET_ROW},1))`, PCT, C.green, true, C.dgreen);
  lblCell(sh, r, 4, "Task Subsidy 按此比例发放（最高100%）", C.lgray, false, C.black, 9, "left", 4);
  sh.setRowHeight(r, 22); const RATE_ROW = r; r+=2;

  // ── B. GDC Projects
  hdrCell(sh, r, 2, "📣 B. GDC Project 收入（3.2.4）", C.teal, C.white, true, 11, 6); sh.setRowHeight(r, 24); r++;
  hdrCell(sh, r, 2, "每个GDC Project填一行 → 内容制作30%池自动汇总", C.lteal, C.teal, false, 9, 6); sh.setRowHeight(r, 16); r++;

  ["Project名称","总收入(RM)","内容制作30%","Lijoo PM 25%","讨论方案5%","DLIG收入40%"].forEach((h,i) =>
    hdrCell(sh, r, i+2, h, C.dgold, C.white, true, 9, 1, true));
  sh.setRowHeight(r, 24); r++;

  const GDC_PROJ_START = r;
  for (let i = 0; i < 6; i++) {
    inpCell(sh, r, 2, i === 0 ? "Project 1" : "");
    inpCell(sh, r, 3, 0, RM);
    frmCell(sh, r, 4, `=C${r}*0.30`, RM, C.lgray);
    frmCell(sh, r, 5, `=C${r}*0.25`, RM, C.lgray);
    frmCell(sh, r, 6, `=C${r}*0.05`, RM, C.lgray);
    frmCell(sh, r, 7, `=C${r}*0.40`, RM, C.lgray);
    sh.setRowHeight(r, 20); r++;
  }
  const GDC_PROJ_END = r - 1;

  lblCell(sh, r, 2, "内容制作总池 (30%)", C.mgray, true, C.black, 10, "center", 2);
  frmCell(sh, r, 4, `=SUM(D${GDC_PROJ_START}:D${GDC_PROJ_END})`, RM, C.mgray, true);
  [5,6,7].forEach(c => sh.getRange(r,c).setBackground(C.mgray));
  const GDC_POOL_ROW = r; sh.setRowHeight(r, 22); r+=2;

  // ── C. GDC 系数分配
  hdrCell(sh, r, 2, "📊 C. GDC 内容分配（按系数占比）", C.teal, C.white, true, 11, 6); sh.setRowHeight(r, 24); r++;
  hdrCell(sh, r, 2, "系数从 Task Log H列自动汇总 → 按占比分 30% 池", C.lteal, C.teal, false, 9, 6); sh.setRowHeight(r, 16); r++;

  ["负责人","GDC总系数","占比%","GDC应得(RM)","",""].forEach((h,i) =>
    hdrCell(sh, r, i+2, h, C.dgold, C.white, true, 9));
  sh.setRowHeight(r, 22); r++;

  const GDC_ROWS = {};
  MEMBERS.forEach(m => {
    lblCell(sh, r, 2, m, C.lteal, true, C.black, 11, "center");
    frmCell(sh, r, 3,
      `=SUMIFS('${logName}'!H${DS}:H${DE},'${logName}'!D${DS}:D${DE},"${m}",'${logName}'!H${DS}:H${DE},"<>")`,
      "0.0", C.lgray);
    sh.getRange(r,4).setBackground(C.lgray).setNumberFormat(PCT).setHorizontalAlignment("center").setVerticalAlignment("middle");
    sh.getRange(r,5).setBackground(C.yellow).setFontColor(C.doran).setFontWeight("bold").setNumberFormat(RM).setHorizontalAlignment("center").setVerticalAlignment("middle");
    [6,7].forEach(c => sh.getRange(r,c).setBackground(C.lgray));
    GDC_ROWS[m] = r; sh.setRowHeight(r, 24); r++;
  });

  lblCell(sh, r, 2, "团队总系数", C.mgray, true, C.black, 10, "center", 2);
  const coefRowList = MEMBERS.map(m => GDC_ROWS[m]);
  frmCell(sh, r, 4, `=SUM(C${coefRowList[0]}:C${coefRowList[coefRowList.length-1]})`, "0.0", C.mgray, true);
  [5,6,7].forEach(c => sh.getRange(r,c).setBackground(C.mgray));
  const TEAM_COEF_ROW = r; sh.setRowHeight(r, 20); r++;

  MEMBERS.forEach(m => {
    const pr = GDC_ROWS[m];
    sh.getRange(pr, 4).setFormula(`=IF(C${TEAM_COEF_ROW}=0,0,C${pr}/C${TEAM_COEF_ROW})`);
    sh.getRange(pr, 5).setFormula(`=D${pr}*D${GDC_POOL_ROW}`);
  });
  r+=2;

  // ── D. DLIG Task Subsidy
  hdrCell(sh, r, 2, "🛠 D. DLIG Task Subsidy（3.2.2）", C.dpurp, C.white, true, 11, 6); sh.setRowHeight(r, 24); r++;
  hdrCell(sh, r, 2, "从 Task Log J列自动汇总 × Sales达成率", C.purple, C.dpurp, false, 9, 6); sh.setRowHeight(r, 16); r++;

  ["负责人","Task小计(RM)","达成率","实际发放(RM)","",""].forEach((h,i) =>
    hdrCell(sh, r, i+2, h, C.dgold, C.white, true, 9));
  sh.setRowHeight(r, 22); r++;

  const DLIG_ROWS = {};
  MEMBERS.forEach(m => {
    lblCell(sh, r, 2, m, C.purple, true, C.black, 11, "center");
    frmCell(sh, r, 3,
      `=SUMIF('${logName}'!D${DS}:D${DE},"${m}",'${logName}'!J${DS}:J${DE})`,
      RM, C.lgray);
    frmCell(sh, r, 4, `=C${RATE_ROW}`, PCT, C.green, false, C.dgreen);
    frmCell(sh, r, 5, `=C${r}*D${r}`, RM, C.yellow, true, C.doran);
    [6,7].forEach(c => sh.getRange(r,c).setBackground(C.lgray));
    DLIG_ROWS[m] = r; sh.setRowHeight(r, 24); r++;
  });

  lblCell(sh, r, 2, "合计", C.mgray, true, C.black, 10, "center");
  const dligRowList = MEMBERS.map(m => DLIG_ROWS[m]);
  frmCell(sh, r, 3, `=SUM(C${dligRowList[0]}:C${dligRowList[dligRowList.length-1]})`, RM, C.mgray, true);
  sh.getRange(r,4).setBackground(C.mgray);
  frmCell(sh, r, 5, `=SUM(E${dligRowList[0]}:E${dligRowList[dligRowList.length-1]})`, RM, C.mgray, true);
  [6,7].forEach(c => sh.getRange(r,c).setBackground(C.mgray));
  sh.setRowHeight(r, 22); r+=2;

  // ── E. 总发放汇总
  hdrCell(sh, r, 2, "✅ E. 总发放汇总（给财务）", C.dgreen, C.white, true, 12, 6); sh.setRowHeight(r, 26); r++;
  ["负责人","GDC应得(RM)","DLIG Task应得(RM)","本月合计(RM)","备注",""].forEach((h,i) =>
    hdrCell(sh, r, i+2, h, C.dgold, C.white, true, 9));
  sh.setRowHeight(r, 22); r++;

  const SUM_ROWS = {};
  MEMBERS.forEach(m => {
    lblCell(sh, r, 2, m, C.green, true, C.black, 12, "center");
    frmCell(sh, r, 3, `=E${GDC_ROWS[m]}`,  RM, C.lgray);
    frmCell(sh, r, 4, `=E${DLIG_ROWS[m]}`, RM, C.lgray);
    frmCell(sh, r, 5, `=C${r}+D${r}`,      RM, C.yellow, true, C.doran);
    inpCell(sh, r, 6, "", null, C.input);
    sh.getRange(r,7).setBackground(C.lgray);
    SUM_ROWS[m] = r; sh.setRowHeight(r, 26); r++;
  });

  const sumRowList = MEMBERS.map(m => SUM_ROWS[m]);
  lblCell(sh, r, 2, "总计 TOTAL", C.mgray, true, C.black, 12, "center");
  frmCell(sh, r, 3, `=SUM(C${sumRowList[0]}:C${sumRowList[sumRowList.length-1]})`, RM, C.mgray, true);
  frmCell(sh, r, 4, `=SUM(D${sumRowList[0]}:D${sumRowList[sumRowList.length-1]})`, RM, C.mgray, true);
  frmCell(sh, r, 5, `=SUM(E${sumRowList[0]}:E${sumRowList[sumRowList.length-1]})`, RM, C.mgray, true, C.dgreen);
  [6,7].forEach(c => sh.getRange(r,c).setBackground(C.mgray));
  sh.setRowHeight(r, 26);
}

// ─── REFERENCE SHEET ────────────────────────────────────────────
function buildReference(sh) {
  sh.clearContents(); sh.clearFormats();
  [20,200,100,100,200].forEach((w,i) => sh.setColumnWidth(i+1, w));

  let r = 1;
  hdrCell(sh, r, 2, "📋 DLIG 单价 & 系数 参考表", C.dblue, C.white, true, 13, 4); sh.setRowHeight(r, 30); r+=2;

  // DLIG tasks
  hdrCell(sh, r, 2, "🛠 DLIG Task Subsidy 单价（3.2.2）", C.dpurp, C.white, true, 11, 4); sh.setRowHeight(r, 22); r++;
  ["Task类型","单价(RM)","类别Key",""].forEach((h,i) => hdrCell(sh, r, i+2, h, C.dgold, C.white, true, 9));
  sh.setRowHeight(r, 20); r++;

  const DLIG = [
    ["视频文案（2篇）",5,"DLIG 视频文案"],["信息文案模版（2篇）",5,"DLIG 文案模版"],
    ["制作Post图+文案（1套）",15,"DLIG Post图文"],["制作海报（1张）",15,"DLIG 海报"],
    ["视频脚本（1支）",5,"DLIG 视频脚本"],["剪辑短视频-一键成片",20,"DLIG 一键成片"],
    ["剪辑短视频（≤1min）",50,"DLIG 短视频剪"],["剪辑YouTube（≤25min）",100,"DLIG YouTube剪"],
    ["Landing Page文案",50,"DLIG LP文案"],["Landing Page制作",150,"DLIG LP制作"],
    ["线下活动带领（半天）",50,"DLIG 活动带领"],["副村长-带盘（半天）",30,"DLIG 副村长带盘"],
    ["Crew-无带盘（半天）",20,"DLIG Crew"],
  ];
  DLIG.forEach(([name, price, key]) => {
    lblCell(sh, r, 2, name, C.purple, false, C.black, 10, "left", 1, true);
    lblCell(sh, r, 3, `RM ${price}`, C.purple, true, C.black, 10, "center");
    lblCell(sh, r, 4, key, C.lgray, false, C.black, 9, "center");
    sh.getRange(r,5).setBackground(C.lgray);
    sh.setRowHeight(r, 20); r++;
  });

  r++;
  hdrCell(sh, r, 2, "📣 GDC 类别系数（3.2.4）", C.teal, C.white, true, 11, 4); sh.setRowHeight(r, 22); r++;
  ["GDC类别","系数","说明",""].forEach((h,i) => hdrCell(sh, r, i+2, h, C.dgold, C.white, true, 9));
  sh.setRowHeight(r, 20); r++;

  const GDC = [
    ["GDC 文案书写",2,"视频文案 ×2"],["GDC 视频文案",2,"视频文案 ×2"],
    ["GDC 图文Post",3,"图文文案 ×3"],["GDC 短视频脚本",1,"脚本 ×1"],
    ["GDC 短视频剪",4,"剪辑短视频 ×4"],["GDC job",0,"PM/对接，按25%+5%另算"],
  ];
  GDC.forEach(([cat, coef, note]) => {
    lblCell(sh, r, 2, cat, C.lteal, true, C.black, 10, "left");
    lblCell(sh, r, 3, coef > 0 ? `×${coef}` : "另算", C.lteal, true, C.teal, 12, "center");
    lblCell(sh, r, 4, note, C.lgray, false, C.black, 9, "left", 2);
    sh.setRowHeight(r, 20); r++;
  });

  r++;
  hdrCell(sh, r, 2, "⚡ 效率系数（Task Log G列填这个）", C.dblue, C.white, true, 11, 4); sh.setRowHeight(r, 22); r++;
  [[1.5,"提前完成","≥2天前完成"],[1.0,"准时完成","按时完成"],[0.5,"延迟完成","3天内延迟"],[0,"未完成","超过3天/未交付"]].forEach(([coef,label,cond]) => {
    lblCell(sh, r, 2, label, C.lblue, true, C.black, 10, "center");
    lblCell(sh, r, 3, `×${coef}`, C.lblue, true, C.dblue, 12, "center");
    lblCell(sh, r, 4, cond, C.lgray, false, C.black, 9, "left", 2);
    sh.setRowHeight(r, 20); r++;
  });
}

// ═══════════════════════════════════════════════════════════════
// END DLIG Task Tracker
// ═══════════════════════════════════════════════════════════════

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
    if (type === 'logs')              return respond(readLogs());
    if (type === 'decisions')         return respond(readDecisions());
    if (type === 'special_meetings')  return respond(readSpecialMeetings());
    if (type === 'attendance')        return respond(readAttendance());
    if (type === 'sb_records')        return respond(readSbRecords());
    if (type === 'au_records')        return respond(readAuRecords());
    if (type === 'xd_roles')          return respond(readXdRoles());
    if (type === 'inventory')         return respond(readInventory());
    if (type === 'bs_all')            return respond(readBSAll());
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
    if (type === 'events')   return respond(writeEventsTab(Array.isArray(data) ? data : []));
    if (type === 'rotation') return respond(writeRotationTab(Array.isArray(data) ? data : []));
    if (type === 'logs')             return respond(writeLogs(Array.isArray(data) ? data : []));
    if (type === 'decisions')        return respond(writeDecisions(Array.isArray(data) ? data : []));
    if (type === 'special_meetings') return respond(writeSpecialMeetings(Array.isArray(data) ? data : []));
    if (type === 'attendance')       return respond(writeAttendance(data));
    if (type === 'sb_records')       return respond(writeSbRecords(Array.isArray(data) ? data : []));
    if (type === 'au_records')       return respond(writeAuRecords(Array.isArray(data) ? data : []));
    if (type === 'xd_roles')         return respond(writeXdRoles(data||{}));
    if (type === 'bs_rotation')      return respond(writeBSRotation(Array.isArray(data)?data:[]));
    if (type === 'pay') {
      return respond(writePaySheet(data.tasks, data.mkt, data.gdc));
    }
    return respond({ error: 'Unknown sheet: ' + type });
  } catch(err) {
    return respond({ error: err.message });
  }
}
