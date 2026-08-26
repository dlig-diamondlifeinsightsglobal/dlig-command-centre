/* DLIG HQ Command Centre 2.0
 * Adds a focused Founder dashboard while preserving DLIG 1.0 pages as history.
 */
(function () {
  const HQ_VERSION = '2.0';
  const STORE = 'dlig_hq2_';
  const API = typeof SHEETS_API_URL !== 'undefined' ? SHEETS_API_URL : '';
  const TODAY = '2026-07-25';

  const DEFAULT_CATEGORIES = [
    'AI Poster + 文案',
    'AI Landing Page + Form + Payment + Auto Message Setup',
    'Social Media Sales AI Content - FB、小红书（15 posts + 广告设置与监管）',
    'Customer Service - WA、FB、小红书',
    'Admin & Account',
    'Event Space Rental Commission',
    'Sales Commission',
    'Event Crew'
  ];

  const ENDPOINTS = {
    revenue: 'hq_revenue',
    monthly: 'hq_monthly',
    products: 'hq_products',
    pipeline: 'hq_pipeline',
    priorities: 'hq_priorities',
    daily: 'hq_daily',
    success: 'hq_success',
    events: 'hq_events',
    operations: 'hq_operations',
    ideas: 'hq_ideas',
    links: 'hq_links',
    tasks: 'hq_tasks',
    categories: 'hq_categories'
  };

  const DEFAULT_STATE = {
    revenue: {
      target: 50000,
      collected: 796,
      awaiting: 1194,
      pipeline: 1194,
      period: '2026-07-01 至 2026-12-31',
      source: 'manual',
      last_updated: '2026-07-25'
    },
    monthly: [
      { id: 'month-jul', month: 'Jul 2026', target: 0, actual: 796, note: 'Pioneer已付款4人' },
      { id: 'month-aug', month: 'Aug 2026', target: 0, actual: 0, note: '' },
      { id: 'month-sep', month: 'Sep 2026', target: 0, actual: 0, note: '' },
      { id: 'month-oct', month: 'Oct 2026', target: 0, actual: 0, note: 'Bootcamp第二期开始' },
      { id: 'month-nov', month: 'Nov 2026', target: 0, actual: 0, note: '学生Camp计划窗口' },
      { id: 'month-dec', month: 'Dec 2026', target: 0, actual: 0, note: '' }
    ],
    products: [
      { id: 'pioneer', product: 'Money-Life Bootcamp', cohort: '第一期 Pioneer', price: 199, min: 12, max: 19, enrolled: 13, paid: 4, transferred: 3, awaiting: 6, start: '2026-08-06', target_revenue: 0, actual_revenue: 796, status: '交付中' },
      { id: 'bootcamp2', product: 'Money-Life Bootcamp', cohort: '第二期', price: 488, min: 20, max: 50, enrolled: 0, paid: 0, transferred: 0, awaiting: 0, start: '2026-10-20', target_revenue: 9760, actual_revenue: 0, status: '准备销售' },
      { id: 'bootcamp3', product: 'Money-Life Bootcamp', cohort: '第三期', price: 538, min: 40, max: 80, enrolled: 0, paid: 0, transferred: 0, awaiting: 0, start: '', target_revenue: 21520, actual_revenue: 0, status: '规划中' },
      { id: 'studentcamp', product: '学生 Camp', cohort: '期数待定', price: 288, min: 50, max: 0, enrolled: 0, paid: 0, transferred: 0, awaiting: 0, start: '2026-11-01', target_revenue: 14400, actual_revenue: 0, status: '产品筹备' },
      { id: 'boss', product: '小老板公司成长讲座', cohort: '期数待定', price: 188, min: 30, max: 0, enrolled: 0, paid: 0, transferred: 0, awaiting: 0, start: '', target_revenue: 5640, actual_revenue: 0, status: '产品筹备' },
      { id: 'workshop', product: 'Money-Life Workshop', cohort: 'Physical Workshop', price: 128, min: 0, max: 0, enrolled: 0, paid: 0, transferred: 0, awaiting: 0, start: '2026-08-13', target_revenue: 0, actual_revenue: 0, status: '每月目标3场' }
    ],
    pipeline: [
      { id: 'lead-paid', group: 'Pioneer 已付款', stage: 'Paid / Enrolled', count: 4, unit_value: 199, expected_value: 796, product: 'Money-Life Bootcamp 第一期', source: 'DLIG member group / 小红书', next_followup: '', note: '已付款', last_updated: '2026-07-25' },
      { id: 'lead-transfer', group: '旧产品服务转入', stage: 'Existing Customer / Next Offer', count: 3, unit_value: 0, expected_value: 0, product: 'Money-Life Bootcamp 第一期', source: 'DLIG 1.0 客户', next_followup: '', note: '无需再次付款；未来可跟进 Alumni Offer', last_updated: '2026-07-25' },
      { id: 'lead-awaiting', group: '已确认、等待付款', stage: 'Confirmed Awaiting Payment', count: 6, unit_value: 199, expected_value: 1194, product: 'Money-Life Bootcamp 第一期', source: 'DLIG member group / 小红书', next_followup: '', note: '其中2位申请8月付款', last_updated: '2026-07-25' },
      { id: 'lead-interested', group: '有兴趣、等待回复', stage: 'Interested', count: 6, unit_value: 199, expected_value: 1194, product: 'Money-Life Bootcamp 第一期', source: 'DLIG member group / 小红书', next_followup: '', note: '需要安排下次跟进日期', last_updated: '2026-07-25' }
    ],
    priorities: [
      { id: 'priority1', outcome: '收齐第一期6位已确认学员的付款安排', owner: 'Founder', status: '进行中', deadline: '2026-08-05', done_definition: '6位都有明确付款状态；2位8月付款者已有确认日期' },
      { id: 'priority2', outcome: '跟进6位有兴趣名单并取得明确回复', owner: 'Founder', status: '待开始', deadline: '2026-07-31', done_definition: '每位都有下一步、日期或明确 Not Now' },
      { id: 'priority3', outcome: '完成第一期 Bootcamp 首课交付准备', owner: 'Founder', status: '进行中', deadline: '2026-08-05', done_definition: '课程内容、Zoom、教材、付款名单与回播流程均检查完成' }
    ],
    daily: [
      { id: 'daily-ai', task: '整理6位兴趣名单的个性化跟进讯息草稿', owner: 'AI', status: '待开始', deadline: '2026-07-27', done_definition: '每位名单都有可直接发送的讯息草稿' },
      { id: 'daily-partner', task: '核对第一期收款与8月付款安排', owner: 'Partner', status: '待开始', deadline: '2026-07-29', done_definition: '已付款、待付款、转入客户三组数字一致' },
      { id: 'daily-is', task: '检查 InfiniteSales 的标签与Follow-up日期', owner: 'InfiniteSales', status: '待开始', deadline: '2026-07-29', done_definition: '每位Lead有正确阶段与下次跟进日期' },
      { id: 'daily-codex', task: '维护 HQ Dashboard 与资料接口', owner: 'Codex', status: '进行中', deadline: '2026-07-31', done_definition: '主要数据可更新、手机可用、README完成' }
    ],
    success: [
      {
        id: 'success-pioneer',
        product: 'Money-Life Bootcamp 第一期 Pioneer',
        promised_result: '6周建立属于自己的 Money-Life：看清真实财务全貌、计算真实时薪、找出消费模式、算出自由数字、建立财富系统；附 Money-Life Templates。',
        next_delivery: '2026-08-06 10:00-12:00',
        feedback: '',
        problem: '',
        improvement: '',
        preparation: '待更新',
        risk: '6位已确认学员仍待付款；6位兴趣名单等待回复',
        owner: 'Founder',
        deadline: '2026-08-05',
        last_updated: '2026-07-25'
      }
    ],
    events: [
      { id: 'ev-bc1-1', date: '2026-08-06', time: '10:00-12:00', name: 'Money-Life Bootcamp Trial 第一期 · Week 1', type: 'Product Class', mode: 'Online', location: 'Online', owner: 'Founder' },
      { id: 'ev-bc1-2', date: '2026-08-13', time: '10:00-12:00', name: 'Money-Life Bootcamp Trial 第一期 · Week 2', type: 'Product Class', mode: 'Online', location: 'Online', owner: 'Founder' },
      { id: 'ev-workshop', date: '2026-08-13', time: '下午（待定）', name: 'Physical Money-Life Workshop 开始', type: 'Workshop', mode: 'Physical', location: 'DLIG Office', owner: 'Founder' },
      { id: 'ev-bc1-3', date: '2026-08-20', time: '10:00-12:00', name: 'Money-Life Bootcamp Trial 第一期 · Week 3', type: 'Product Class', mode: 'Online', location: 'Online', owner: 'Founder' },
      { id: 'ev-bc1-4', date: '2026-08-27', time: '10:00-12:00', name: 'Money-Life Bootcamp Trial 第一期 · Week 4', type: 'Product Class', mode: 'Online', location: 'Online', owner: 'Founder' },
      { id: 'ev-bc1-5', date: '2026-09-03', time: '10:00-12:00', name: 'Money-Life Bootcamp Trial 第一期 · Week 5', type: 'Product Class', mode: 'Online', location: 'Online', owner: 'Founder' },
      { id: 'ev-bc1-6', date: '2026-09-10', time: '10:00-12:00', name: 'Money-Life Bootcamp Trial 第一期 · Week 6', type: 'Product Class', mode: 'Online', location: 'Online', owner: 'Founder' },
      { id: 'ev-free-talk', date: '2026-08-19', time: '启动日', name: 'Bootcamp Free Talk Marketing 规划开始', type: 'Planning', mode: 'Online', location: 'Online', owner: 'Founder' },
      { id: 'ev-bc2', date: '2026-10-20', time: '19:45-21:45', name: 'Money-Life Bootcamp 正式版第二期', type: 'Product Class', mode: 'Online', location: 'Online', owner: 'Founder' },
      { id: 'ev-student', date: '2026-11-01', time: '日期待定', name: '学生 Camp 销售与交付窗口', type: 'Planning', mode: 'TBD', location: 'DLIG Office / TBD', owner: 'Founder' }
    ],
    operations: [
      { id: 'op-recording', task: '上传第一期课程回播并检查客户权限', owner: 'Partner', status: '待开始', deadline: '2026-08-07', done_definition: '回播上传、命名正确，并用客户账号测试可观看' },
      { id: 'op-course', task: '建立6周课程交付清单', owner: 'Founder', status: '进行中', deadline: '2026-08-05', done_definition: '每周内容、模板、提醒、出席和反馈流程都有负责人' },
      { id: 'op-receipt', task: '每周整理 Receipt Claim', owner: 'Partner', status: '循环任务', deadline: '', done_definition: '当周收据已上传并可追溯' }
    ],
    ideas: [
      { id: 'idea-alumni', idea: 'Alumni Price - 第二期开始', reason: '让第一期学员可继续参与或推荐，同时区别新客价格', potential_value: '复购、转介绍与社群延续', revisit_date: '2026-09-15', status: '待评估' },
      { id: 'idea-club', idea: 'Money-Life Club（Subscription）', reason: '取代单纯复训；让 Money System 持续实践并成为生活方式', potential_value: '每月 Review Gathering、新主题、新工具、Recording Library、社群与陪伴', revisit_date: '2026-11-15', status: 'Parking' }
    ],
    links: [
      { id: 'link-is', name: 'InfiniteSales System', url: 'https://app.infinitesales.ai/v2/location/qsvfYF9CIZ99cPzvLL4a/dashboard', icon: '∞', note: 'CRM 与跟进' },
      { id: 'link-canva', name: 'DLIG 2.0 资料', url: 'https://canva.link/nn8bd5bxdaa2kn5', icon: '🎨', note: 'Canva 2.0' },
      { id: 'link-receipts', name: 'Receipt Claim 存档', url: 'https://drive.google.com/drive/folders/1ex9M8svLswmXVImmB08sZlB9hHkF84J1?usp=drive_link', icon: '🧾', note: 'Google Drive' },
      { id: 'link-classroom', name: '客户回播 / 课室', url: 'https://diamondlifeinsightsglobal.app.clientclub.net/', icon: '▶', note: '客户课程权限' },
      { id: 'link-sales', name: 'Sales & Marketing Sheet', url: 'https://docs.google.com/spreadsheets/d/1F0Ss1MuAwVRkVfWch2wepx2SvMXG3BoIBfVN0v9ye9M/edit', icon: '📊', note: '现有销售资料' },
      { id: 'link-task', name: 'Task & Events Sheet', url: 'https://docs.google.com/spreadsheets/d/1A3g_WPDU-R4zU8gj8gGHu885z5bk4lTBBRnCfDsstDI/edit', icon: '✅', note: '任务与活动资料' }
    ],
    tasks: [],
    categories: DEFAULT_CATEGORIES.map((name, index) => ({ id: 'cat-' + index, name })),
    meta: {}
  };

  const STAGES = [
    'New Leads',
    'Contacted',
    'Engaged',
    'Interested',
    'Decision Pending',
    'Confirmed Awaiting Payment',
    'Paid / Enrolled',
    'Not Now',
    'Lost',
    'Existing Customer / Next Offer'
  ];

  const OWNER_OPTIONS = ['Founder', 'Partner', 'AI', 'InfiniteSales', 'Codex'];
  const STATUS_OPTIONS = ['待开始', '进行中', '等待中', '已完成', '循环任务'];
  let state = loadState();
  let taskOwnerFilter = '全部';
  let taskStatusFilter = '全部';
  let pipelineStageFilter = '全部';

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function loadState() {
    const result = clone(DEFAULT_STATE);
    Object.keys(DEFAULT_STATE).forEach(key => {
      try {
        const saved = localStorage.getItem(STORE + key);
        if (saved) result[key] = JSON.parse(saved);
      } catch (error) {}
    });
    return result;
  }

  function esc(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function rm(value) {
    const number = Number(value) || 0;
    return 'RM' + number.toLocaleString('en-MY', { maximumFractionDigits: 0 });
  }

  function fmtDate(value) {
    if (!value) return '待定';
    const date = new Date(value + (String(value).length === 10 ? 'T00:00:00' : ''));
    if (Number.isNaN(date.getTime())) return value;
    return date.toLocaleDateString('zh-MY', { day: 'numeric', month: 'short', year: 'numeric' });
  }

  function isoNow() {
    return new Date().toISOString();
  }

  function updatedLabel(key) {
    const meta = state.meta[key] || {};
    const date = meta.updated || state.revenue.last_updated || TODAY;
    const source = meta.source === 'sheet' ? 'Google Sheet 已同步' : '手动资料';
    return `<span class="hq-source">${source} · ${esc(String(date).slice(0, 16).replace('T', ' '))}</span>`;
  }

  function persist(key) {
    const updated = isoNow();
    state.meta[key] = { updated, source: 'manual' };
    try {
      localStorage.setItem(STORE + key, JSON.stringify(state[key]));
      localStorage.setItem(STORE + 'meta', JSON.stringify(state.meta));
    } catch (error) {}
    syncModule(key);
  }

  function syncModule(key) {
    if (!API || !ENDPOINTS[key]) return;
    const value = Array.isArray(state[key]) ? state[key] : [state[key]];
    fetch(`${API}?sheet=${ENDPOINTS[key]}`, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(value)
    }).then(response => response.json()).then(data => {
      if (data && (data.ok || data.count >= 0 || data.rows >= 0)) {
        state.meta[key] = { updated: isoNow(), source: 'sheet' };
        try { localStorage.setItem(STORE + 'meta', JSON.stringify(state.meta)); } catch (error) {}
        renderCurrent();
      }
    }).catch(() => {});
  }

  async function loadHQData() {
    if (!API) return;
    await Promise.all(Object.keys(ENDPOINTS).map(async key => {
      try {
        const response = await fetch(`${API}?sheet=${ENDPOINTS[key]}`);
        const rows = await response.json();
        if (!Array.isArray(rows) || !rows.length) return;
        state[key] = Array.isArray(DEFAULT_STATE[key]) ? rows : rows[0];
        state.meta[key] = { updated: isoNow(), source: 'sheet' };
        localStorage.setItem(STORE + key, JSON.stringify(state[key]));
      } catch (error) {}
    }));
    try { localStorage.setItem(STORE + 'meta', JSON.stringify(state.meta)); } catch (error) {}
    renderAllHQ();
  }

  function ensurePage(id) {
    let page = document.getElementById('p-' + id);
    if (!page) {
      page = document.createElement('div');
      page.className = 'page';
      page.id = 'p-' + id;
      document.querySelector('.main').appendChild(page);
    }
    return page;
  }

  function pageHead(kicker, title, desc, key) {
    return `<div class="hq-page-head">
      <div>
        <div class="hq-page-kicker">${esc(kicker)}</div>
        <div class="hq-page-title">${esc(title)}</div>
        <div class="hq-page-desc">${esc(desc)}</div>
      </div>
      ${updatedLabel(key)}
    </div>`;
  }

  function setNavigation() {
    const nav = document.querySelector('.nav');
    nav.innerHTML = `
      <div class="hq-nav-section">Founder View</div>
      ${navItem('overview', '🏠', 'HQ 总览')}
      ${navItem('hq-revenue', '🎯', 'Revenue')}
      ${navItem('hq-pipeline', '🤝', 'Sales Pipeline')}
      <div class="hq-nav-section">Execution</div>
      ${navItem('hq-daily', '☀️', 'Daily Command')}
      ${navItem('tasks', '✅', 'Task Board')}
      ${navItem('calendar', '📅', 'Calendar')}
      ${navItem('hq-success', '🌱', 'Customer Success')}
      ${navItem('hq-operations', '⚙️', 'Operations')}
      <div class="hq-nav-section">原有业务资料</div>
      ${legacyNavItem('products', '🎁', '产品 & 价格')}
      ${legacyNavItem('sales', '💰', 'Sales & Marketing')}
      ${legacyNavItem('admin', '🗂', 'Admin & Account')}
      <div class="hq-nav-section">Focus Control</div>
      ${navItem('hq-ideas', '💡', 'Idea Parking Lot')}
      ${navItem('files', '🗂', '文件索引 / 1.0 记录')}
    `;
  }

  function navItem(id, icon, label) {
    return `<div class="ni" data-hq-nav="${id}" onclick="hqGo('${id}')"><span class="ni-icon">${icon}</span>${label}</div>`;
  }

  function legacyNavItem(id, icon, label) {
    return `<div class="ni" data-hq-legacy="${id}" onclick="hqOpenLegacy('${id}','${esc(label)}')"><span class="ni-icon">${icon}</span>${label}</div>`;
  }

  function setTopbar() {
    const topbar = document.querySelector('.topbar');
    if (!topbar.querySelector('.hq-mobile-menu')) {
      const button = document.createElement('button');
      button.className = 'hq-mobile-menu';
      button.type = 'button';
      button.setAttribute('aria-label', '打开菜单');
      button.textContent = '☰';
      button.onclick = () => document.body.classList.toggle('sidebar-open');
      topbar.insertBefore(button, topbar.firstChild);
    }
    document.querySelector('.tb-sub').textContent = 'July-Dec 2026 · Revenue Target RM50,000';
    document.title = 'DLIG HQ Command Centre 2.0';
  }

  function hqGo(id, skipHash) {
    document.querySelectorAll('.page').forEach(page => page.classList.remove('on'));
    document.querySelectorAll('[data-hq-nav]').forEach(item => item.classList.toggle('on', item.dataset.hqNav === id));
    document.querySelectorAll('[data-hq-legacy]').forEach(item => item.classList.remove('on'));
    const page = document.getElementById('p-' + id);
    if (page) page.classList.add('on');
    const titleMap = {
      overview: '🏠 DLIG HQ 总览',
      'hq-revenue': '🎯 Revenue Dashboard',
      'hq-pipeline': '🤝 Sales Pipeline',
      'hq-daily': '☀️ Daily Command Centre',
      tasks: '✅ Task Board',
      calendar: '📅 Calendar',
      'hq-success': '🌱 Customer Success',
      'hq-operations': '⚙️ Operations',
      'hq-ideas': '💡 Idea Parking Lot',
      files: '🗂 文件索引'
    };
    document.getElementById('pg-title').textContent = titleMap[id] || id;
    document.getElementById('tb-pills').style.display = 'none';
    document.body.classList.remove('sidebar-open');
    if (!skipHash) history.replaceState(null, '', '#' + id);
    renderById(id);
    window.scrollTo(0, 0);
  }

  function openLegacy(id, title) {
    document.querySelectorAll('.page').forEach(page => page.classList.remove('on'));
    document.querySelectorAll('[data-hq-nav]').forEach(item => item.classList.remove('on'));
    document.querySelectorAll('[data-hq-legacy]').forEach(item => item.classList.toggle('on', item.dataset.hqLegacy === id));
    const page = document.getElementById('p-' + id);
    if (!page) return;
    page.classList.add('on');
    let banner = page.querySelector('.hq-legacy-banner');
    if (!banner) {
      banner = document.createElement('div');
      banner.className = 'hq-legacy-banner';
      banner.innerHTML = `<span>📦 这是原有的 DLIG 1.0 业务资料，内容与功能继续保留；不会自动成为 HQ 2.0 今日优先事项。</span><button class="hq-btn secondary" onclick="hqGo('files')">打开文件索引</button>`;
      page.insertBefore(banner, page.firstChild);
    }
    document.getElementById('pg-title').textContent = '📦 1.0记录 · ' + title;
    history.replaceState(null, '', '#legacy-' + id);
    window.scrollTo(0, 0);
  }

  function renderHQOverview() {
    const page = document.getElementById('p-overview');
    if (!page) return;
    const revenue = state.revenue;
    const shortfall = Math.max(0, revenue.target - revenue.collected);
    const committedShortfall = Math.max(0, revenue.target - revenue.collected - revenue.awaiting);
    const pct = revenue.target ? Math.min(100, revenue.collected / revenue.target * 100) : 0;
    const priorities = state.priorities.slice(0, 3);
    const nextEvent = state.events.filter(event => event.date >= TODAY).sort((a, b) => a.date.localeCompare(b.date))[0];
    const awaitingGroup = state.pipeline.find(row => row.stage === 'Confirmed Awaiting Payment');
    const interestedGroup = state.pipeline.find(row => row.stage === 'Interested');
    const overdueTasks = state.tasks.filter(task => !task.completed && task.deadline && task.deadline < TODAY).length;
    const followupRisk = state.pipeline.filter(row => ['Interested', 'Decision Pending', 'Confirmed Awaiting Payment'].includes(row.stage) && !row.next_followup).length;

    page.innerHTML = `
      ${pageHead('DLIG HQ 2.0', 'RM50,000 Revenue Mission', 'Founder 每天只需要看这一页：钱、最接近购买的人、下一次交付、今天三大成果，以及风险。', 'revenue')}
      <div class="hq-metric-grid">
        ${metric('已收款', rm(revenue.collected), 'Pioneer 4人 × RM199', 'emphasis')}
        ${metric('确认待付款', rm(revenue.awaiting), `${awaitingGroup ? awaitingGroup.count : 0}人 · 其中2人计划8月付款`, '')}
        ${metric('兴趣 Pipeline', rm(revenue.pipeline), `${interestedGroup ? interestedGroup.count : 0}人等待回复`, '')}
        ${metric('收入缺口', rm(shortfall), `计入待付款后仍差 ${rm(committedShortfall)}`, 'warning')}
        ${metric('Jul-Dec目标', rm(revenue.target), `${pct.toFixed(1)}% 已收款`, '')}
      </div>
      <div class="hq-panel" style="margin-bottom:14px">
        <div class="hq-panel-title"><span>目标进度</span><span>${pct.toFixed(1)}%</span></div>
        <div class="hq-progress"><span style="width:${pct}%"></span></div>
        <div class="hq-note" style="margin-top:8px">只按已实际收款计算。待付款与兴趣名单分别显示，不伪装成已收入。</div>
      </div>
      <div class="hq-grid-2">
        <div class="hq-panel">
          <div class="hq-panel-title"><span>🔥 最接近购买</span><span class="hq-status amber">立即跟进</span></div>
          <div style="font-size:17px;font-weight:800;color:var(--hq-ink)">6位已确认、等待付款</div>
          <div class="hq-note" style="margin-top:7px">预计 ${rm(revenue.awaiting)} · 其中2位申请在8月付款。下一步是为每人补上明确付款日期。</div>
        </div>
        <div class="hq-panel">
          <div class="hq-panel-title"><span>📅 下一次客户交付</span><span class="hq-status blue">${nextEvent ? fmtDate(nextEvent.date) : '待定'}</span></div>
          <div style="font-size:15px;font-weight:800">${nextEvent ? esc(nextEvent.name) : '暂无活动'}</div>
          <div class="hq-note" style="margin-top:7px">${nextEvent ? `${esc(nextEvent.time)} · ${esc(nextEvent.mode)} · ${esc(nextEvent.location)}` : ''}</div>
        </div>
      </div>
      <div class="hq-grid-2">
        <div class="hq-panel">
          <div class="hq-panel-title"><span>☀️ Founder 今日三大成果</span><button class="hq-btn secondary" onclick="hqGo('hq-daily')">更新</button></div>
          ${priorities.map((item, index) => priorityRow(item, index)).join('')}
        </div>
        <div class="hq-panel">
          <div class="hq-panel-title"><span>⚠️ Overdue / At Risk</span><span class="hq-status ${overdueTasks + followupRisk ? 'red' : 'green'}">${overdueTasks + followupRisk}项</span></div>
          <div class="hq-callout"><b>${followupRisk}个 Lead Group 尚未安排下次跟进日期</b><br>第一期另有6位等待付款。请在 Sales Pipeline 补上日期，让提醒真正可用。</div>
          <div class="hq-note" style="margin-top:10px">HQ Task Board 当前逾期：${overdueTasks}。DLIG 1.0 的旧任务不会计入。</div>
        </div>
      </div>
      <div class="hq-grid-3">
        ${miniProductCard(state.products.find(p => p.id === 'pioneer'))}
        ${miniProductCard(state.products.find(p => p.id === 'bootcamp2'))}
        ${miniProductCard(state.products.find(p => p.id === 'bootcamp3'))}
      </div>`;
  }

  function metric(label, value, note, className) {
    return `<div class="hq-metric ${className || ''}">
      <div class="hq-metric-label">${esc(label)}</div>
      <div class="hq-metric-value">${esc(value)}</div>
      <div class="hq-metric-note">${esc(note)}</div>
    </div>`;
  }

  function priorityRow(item, index) {
    return `<div class="hq-priority">
      <div class="hq-priority-num">${index + 1}</div>
      <div><div class="hq-priority-text">${esc(item.outcome)}</div><div class="hq-priority-sub">${esc(item.owner)} · ${fmtDate(item.deadline)}</div></div>
      ${statusBadge(item.status)}
    </div>`;
  }

  function miniProductCard(product) {
    if (!product) return '';
    const denominator = product.min || product.max || 1;
    const pct = Math.min(100, product.enrolled / denominator * 100);
    return `<div class="hq-panel">
      <div class="hq-panel-title"><span>${esc(product.cohort)}</span>${statusBadge(product.status)}</div>
      <div style="font-size:20px;font-weight:800">${product.enrolled}<span style="font-size:11px;color:var(--muted)"> / 最低${product.min}</span></div>
      <div class="hq-progress"><span style="width:${pct}%"></span></div>
      <div class="hq-note" style="margin-top:7px">上限 ${product.max || '待定'} · 当前价 ${rm(product.price)}</div>
    </div>`;
  }

  function statusBadge(status) {
    const value = String(status || '');
    const className = /已完成|已付款|交付中/.test(value) ? 'green'
      : /进行中|准备|确认|循环/.test(value) ? 'amber'
      : /风险|逾期|Lost/.test(value) ? 'red'
      : /规划|等待|待评估/.test(value) ? 'blue' : 'gray';
    return `<span class="hq-status ${className}">${esc(value || '待更新')}</span>`;
  }

  function renderRevenuePage() {
    const page = ensurePage('hq-revenue');
    const futureMinimum = state.products.filter(product => product.id !== 'pioneer').reduce((sum, product) => sum + (Number(product.target_revenue) || 0), 0);
    const actual = state.products.reduce((sum, product) => sum + (Number(product.actual_revenue) || 0), 0);
    const planned = futureMinimum + actual;
    const shortfall = Math.max(0, state.revenue.target - state.revenue.collected);
    page.innerHTML = `
      ${pageHead('Revenue Dashboard', 'July-December 2026', '目标、实际收款、待付款、Pipeline，以及每个产品的最低人数与收入计划。', 'revenue')}
      <div class="hq-metric-grid">
        ${metric('总目标', rm(state.revenue.target), state.revenue.period, 'emphasis')}
        ${metric('已收款', rm(state.revenue.collected), '只计实际到账', '')}
        ${metric('确认待付款', rm(state.revenue.awaiting), '不计入实际收入', '')}
        ${metric('Pipeline', rm(state.revenue.pipeline), '兴趣名单预计值', '')}
        ${metric('Shortfall', rm(shortfall), '目标减已收款', 'warning')}
      </div>
      <div class="hq-form">
        ${field('Jul-Dec目标', 'hq-rev-target', 'number', state.revenue.target)}
        ${field('已收款', 'hq-rev-collected', 'number', state.revenue.collected)}
        ${field('确认待付款', 'hq-rev-awaiting', 'number', state.revenue.awaiting)}
        ${field('兴趣Pipeline', 'hq-rev-pipeline', 'number', state.revenue.pipeline)}
        <button class="hq-btn" onclick="hqSaveRevenue()">保存 Revenue</button>
      </div>
      <div class="hq-callout" style="margin-bottom:12px"><b>最低产品收入计划：${rm(planned)}</b>（未来产品最低计划 ${rm(futureMinimum)} + Pioneer已收 ${rm(actual)}）。第二期使用RM488 Super Early Bird；第三期暂以“至少加RM50”的RM538计划价估算，可随时修改。</div>
      <div class="hq-panel" style="margin-bottom:12px">
        <div class="hq-panel-title"><span>Monthly Sales</span><span class="hq-note">尚未决定的月目标可保留为0，之后再填写</span></div>
        <div class="hq-table-wrap"><table class="hq-table" style="min-width:620px">
          <thead><tr><th>月份</th><th>月目标</th><th>实际收款</th><th>差额</th><th>备注</th></tr></thead>
          <tbody>${state.monthly.map((month, index) => `<tr>
            <td><b>${esc(month.month)}</b></td>
            <td><input type="number" value="${Number(month.target) || 0}" onchange="hqUpdateMonthly(${index},'target',this.value)"></td>
            <td><input type="number" value="${Number(month.actual) || 0}" onchange="hqUpdateMonthly(${index},'actual',this.value)"></td>
            <td>${rm(Math.max(0, (Number(month.target) || 0) - (Number(month.actual) || 0)))}</td>
            <td><input value="${esc(month.note || '')}" onchange="hqUpdateMonthly(${index},'note',this.value)"></td>
          </tr>`).join('')}</tbody>
        </table></div>
      </div>
      <div class="hq-table-wrap">
        <table class="hq-table">
          <thead><tr><th>产品/期数</th><th>价格</th><th>最低人数</th><th>上限</th><th>当前确认</th><th>已付款</th><th>目标收入</th><th>实际收入</th><th>开始</th></tr></thead>
          <tbody>${state.products.map((product, index) => `
            <tr>
              <td><b>${esc(product.product)}</b><br><span class="hq-note">${esc(product.cohort)}</span></td>
              <td><input type="number" value="${Number(product.price) || 0}" onchange="hqUpdateProduct(${index},'price',this.value)"></td>
              <td><input type="number" value="${Number(product.min) || 0}" onchange="hqUpdateProduct(${index},'min',this.value)"></td>
              <td><input type="number" value="${Number(product.max) || 0}" onchange="hqUpdateProduct(${index},'max',this.value)"></td>
              <td><input type="number" value="${Number(product.enrolled) || 0}" onchange="hqUpdateProduct(${index},'enrolled',this.value)"></td>
              <td><input type="number" value="${Number(product.paid) || 0}" onchange="hqUpdateProduct(${index},'paid',this.value)"></td>
              <td><input type="number" value="${Number(product.target_revenue) || 0}" onchange="hqUpdateProduct(${index},'target_revenue',this.value)"></td>
              <td><input type="number" value="${Number(product.actual_revenue) || 0}" onchange="hqUpdateProduct(${index},'actual_revenue',this.value)"></td>
              <td>${fmtDate(product.start)}</td>
            </tr>`).join('')}</tbody>
        </table>
      </div>
      <div class="hq-panel" style="margin-top:12px">
        <div class="hq-panel-title"><span>Bootcamp 第二期价格阶梯</span><span class="hq-status amber">9/9/2026截止</span></div>
        <div class="hq-grid-3" style="margin-bottom:0">
          ${metric('课程价值', 'RM999', 'Regular Value', '')}
          ${metric('正式售价', 'RM668', 'Selling Price', '')}
          ${metric('Super Early Bird', 'RM488', '9/9/2026前', 'emphasis')}
        </div>
        <div class="hq-note">未来每一期计划至少增加RM50，逐步回到Standard Rate；价格仍可在产品表内更新。</div>
      </div>`;
  }

  function field(label, id, type, value, extraClass) {
    return `<div class="hq-field ${extraClass || ''}"><label for="${id}">${esc(label)}</label><input id="${id}" type="${type}" value="${esc(value)}"></div>`;
  }

  function hqSaveRevenue() {
    state.revenue.target = Number(document.getElementById('hq-rev-target').value) || 0;
    state.revenue.collected = Number(document.getElementById('hq-rev-collected').value) || 0;
    state.revenue.awaiting = Number(document.getElementById('hq-rev-awaiting').value) || 0;
    state.revenue.pipeline = Number(document.getElementById('hq-rev-pipeline').value) || 0;
    state.revenue.last_updated = isoNow();
    persist('revenue');
    renderRevenuePage();
    renderHQOverview();
  }

  function hqUpdateProduct(index, key, value) {
    if (!state.products[index]) return;
    state.products[index][key] = Number(value) || 0;
    persist('products');
    renderRevenuePage();
    renderHQOverview();
  }

  function hqUpdateMonthly(index, key, value) {
    if (!state.monthly[index]) return;
    state.monthly[index][key] = key === 'target' || key === 'actual' ? Number(value) || 0 : value;
    persist('monthly');
    renderRevenuePage();
  }

  function renderPipelinePage() {
    const page = ensurePage('hq-pipeline');
    const counts = Object.fromEntries(STAGES.map(stage => [stage, state.pipeline.filter(row => row.stage === stage).reduce((sum, row) => sum + Number(row.count || 0), 0)]));
    const filtered = pipelineStageFilter === '全部' ? state.pipeline : state.pipeline.filter(row => row.stage === pipelineStageFilter);
    const totalValue = state.pipeline.filter(row => !['Paid / Enrolled', 'Lost', 'Not Now'].includes(row.stage)).reduce((sum, row) => sum + Number(row.expected_value || 0), 0);
    page.innerHTML = `
      ${pageHead('Sales Dashboard', 'Lead Groups & Follow-up', 'Phase 1先以Lead Group管理，不需要把客户私人资料写入GitHub。日后可由InfiniteSales API/Webhook替换资料来源。', 'pipeline')}
      <div class="hq-metric-grid">
        ${metric('Active Pipeline', rm(totalValue), '未含已付款、Lost及Not Now', 'emphasis')}
        ${metric('Confirmed Awaiting', counts['Confirmed Awaiting Payment'], rm(state.revenue.awaiting), '')}
        ${metric('Interested', counts.Interested, rm(state.revenue.pipeline), '')}
        ${metric('Paid / Enrolled', counts['Paid / Enrolled'], '另有3位旧产品转入', '')}
        ${metric('缺Follow-up日期', state.pipeline.filter(row => ['Interested', 'Decision Pending', 'Confirmed Awaiting Payment'].includes(row.stage) && !row.next_followup).length, '需要安排', 'warning')}
      </div>
      <div class="hq-filter-row">
        ${['全部'].concat(STAGES).map(stage => `<button class="hq-filter ${pipelineStageFilter === stage ? 'on' : ''}" onclick="hqSetPipelineFilter('${esc(stage)}')">${esc(stage)} · ${stage === '全部' ? state.pipeline.reduce((sum, row) => sum + Number(row.count || 0), 0) : counts[stage]}</button>`).join('')}
      </div>
      <div class="hq-form">
        ${field('Lead Group', 'hq-lead-group', 'text', '', 'wide')}
        <div class="hq-field"><label>Stage</label><select id="hq-lead-stage">${STAGES.map(stage => `<option>${esc(stage)}</option>`).join('')}</select></div>
        ${field('人数', 'hq-lead-count', 'number', 1)}
        ${field('每人预计金额', 'hq-lead-unit', 'number', 0)}
        ${field('产品', 'hq-lead-product', 'text', 'Money-Life Bootcamp')}
        ${field('来源', 'hq-lead-source', 'text', '')}
        ${field('下次跟进', 'hq-lead-followup', 'date', '')}
        <button class="hq-btn" onclick="hqAddPipeline()">加入Pipeline</button>
      </div>
      <div class="hq-table-wrap">
        <table class="hq-table">
          <thead><tr><th>Lead Group</th><th>Stage</th><th>人数</th><th>Expected Value</th><th>产品/来源</th><th>下次跟进</th><th>备注</th><th></th></tr></thead>
          <tbody>${filtered.map(row => {
            const index = state.pipeline.indexOf(row);
            return `<tr>
              <td><b>${esc(row.group)}</b></td>
              <td><select onchange="hqUpdatePipeline(${index},'stage',this.value)">${STAGES.map(stage => `<option ${stage === row.stage ? 'selected' : ''}>${esc(stage)}</option>`).join('')}</select></td>
              <td><input type="number" value="${Number(row.count) || 0}" onchange="hqUpdatePipeline(${index},'count',this.value)"></td>
              <td>${rm(row.expected_value)}</td>
              <td>${esc(row.product)}<br><span class="hq-note">${esc(row.source)}</span></td>
              <td><input type="date" value="${esc(row.next_followup || '')}" onchange="hqUpdatePipeline(${index},'next_followup',this.value)"></td>
              <td><input value="${esc(row.note || '')}" onchange="hqUpdatePipeline(${index},'note',this.value)"></td>
              <td><button class="hq-btn danger" onclick="hqDeletePipeline(${index})">删除</button></td>
            </tr>`;
          }).join('')}</tbody>
        </table>
      </div>`;
  }

  function hqSetPipelineFilter(stage) {
    pipelineStageFilter = stage;
    renderPipelinePage();
  }

  function hqAddPipeline() {
    const group = document.getElementById('hq-lead-group').value.trim();
    if (!group) return;
    const count = Number(document.getElementById('hq-lead-count').value) || 0;
    const unit = Number(document.getElementById('hq-lead-unit').value) || 0;
    state.pipeline.push({
      id: 'lead-' + Date.now(),
      group,
      stage: document.getElementById('hq-lead-stage').value,
      count,
      unit_value: unit,
      expected_value: count * unit,
      product: document.getElementById('hq-lead-product').value,
      source: document.getElementById('hq-lead-source').value,
      next_followup: document.getElementById('hq-lead-followup').value,
      note: '',
      last_updated: isoNow()
    });
    persist('pipeline');
    renderPipelinePage();
    renderHQOverview();
  }

  function hqUpdatePipeline(index, key, value) {
    const row = state.pipeline[index];
    if (!row) return;
    row[key] = key === 'count' || key === 'unit_value' ? Number(value) || 0 : value;
    row.expected_value = (Number(row.count) || 0) * (Number(row.unit_value) || 0);
    row.last_updated = isoNow();
    persist('pipeline');
    renderPipelinePage();
    renderHQOverview();
  }

  function hqDeletePipeline(index) {
    state.pipeline.splice(index, 1);
    persist('pipeline');
    renderPipelinePage();
    renderHQOverview();
  }

  function renderDailyPage() {
    const page = ensurePage('hq-daily');
    page.innerHTML = `
      ${pageHead('Daily Command Centre', 'Founder Top 3 + Delegated Work', '每项工作都要有Owner、状态、Deadline和Definition of Done。Top 3可直接在这里更新。', 'daily')}
      <div class="hq-panel" style="margin-bottom:12px">
        <div class="hq-panel-title"><span>Founder 今日三大成果</span><span class="hq-status blue">最多3项</span></div>
        ${state.priorities.slice(0, 3).map((item, index) => `
          <div class="hq-form" style="margin-bottom:8px">
            ${field('Outcome ' + (index + 1), 'hq-priority-' + index, 'text', item.outcome, 'wide')}
            <div class="hq-field"><label>状态</label><select id="hq-priority-status-${index}">${STATUS_OPTIONS.map(status => `<option ${status === item.status ? 'selected' : ''}>${status}</option>`).join('')}</select></div>
            ${field('Deadline', 'hq-priority-deadline-' + index, 'date', item.deadline)}
            ${field('Definition of Done', 'hq-priority-done-' + index, 'text', item.done_definition, 'wide')}
            <button class="hq-btn" onclick="hqSavePriority(${index})">保存</button>
          </div>`).join('')}
      </div>
      <div class="hq-form">
        ${field('Task', 'hq-daily-task', 'text', '', 'wide')}
        <div class="hq-field"><label>Owner</label><select id="hq-daily-owner">${OWNER_OPTIONS.map(owner => `<option>${owner}</option>`).join('')}</select></div>
        ${field('Deadline', 'hq-daily-deadline', 'date', '')}
        ${field('Definition of Done', 'hq-daily-done', 'text', '', 'wide')}
        <button class="hq-btn" onclick="hqAddDaily()">添加工作</button>
      </div>
      ${editableActionTable(state.daily, 'daily')}`;
  }

  function editableActionTable(rows, key) {
    return `<div class="hq-table-wrap"><table class="hq-table">
      <thead><tr><th>Task</th><th>Owner</th><th>Status</th><th>Deadline</th><th>Definition of Done</th><th></th></tr></thead>
      <tbody>${rows.map((row, index) => `<tr>
        <td><input value="${esc(row.task)}" onchange="hqUpdateAction('${key}',${index},'task',this.value)"></td>
        <td><select onchange="hqUpdateAction('${key}',${index},'owner',this.value)">${OWNER_OPTIONS.map(owner => `<option ${owner === row.owner ? 'selected' : ''}>${owner}</option>`).join('')}</select></td>
        <td><select onchange="hqUpdateAction('${key}',${index},'status',this.value)">${STATUS_OPTIONS.map(status => `<option ${status === row.status ? 'selected' : ''}>${status}</option>`).join('')}</select></td>
        <td><input type="date" value="${esc(row.deadline || '')}" onchange="hqUpdateAction('${key}',${index},'deadline',this.value)"></td>
        <td><input value="${esc(row.done_definition || '')}" onchange="hqUpdateAction('${key}',${index},'done_definition',this.value)"></td>
        <td><button class="hq-btn danger" onclick="hqDeleteAction('${key}',${index})">删除</button></td>
      </tr>`).join('')}</tbody>
    </table></div>`;
  }

  function hqSavePriority(index) {
    const row = state.priorities[index];
    if (!row) return;
    row.outcome = document.getElementById('hq-priority-' + index).value;
    row.status = document.getElementById('hq-priority-status-' + index).value;
    row.deadline = document.getElementById('hq-priority-deadline-' + index).value;
    row.done_definition = document.getElementById('hq-priority-done-' + index).value;
    persist('priorities');
    renderDailyPage();
    renderHQOverview();
  }

  function hqAddDaily() {
    const task = document.getElementById('hq-daily-task').value.trim();
    if (!task) return;
    state.daily.push({
      id: 'daily-' + Date.now(),
      task,
      owner: document.getElementById('hq-daily-owner').value,
      status: '待开始',
      deadline: document.getElementById('hq-daily-deadline').value,
      done_definition: document.getElementById('hq-daily-done').value
    });
    persist('daily');
    renderDailyPage();
  }

  function hqUpdateAction(key, index, fieldName, value) {
    if (!state[key] || !state[key][index]) return;
    state[key][index][fieldName] = value;
    persist(key);
    renderById(key === 'daily' ? 'hq-daily' : 'hq-operations');
  }

  function hqDeleteAction(key, index) {
    state[key].splice(index, 1);
    persist(key);
    renderById(key === 'daily' ? 'hq-daily' : 'hq-operations');
  }

  function renderHQTasks() {
    const page = document.getElementById('p-tasks');
    if (!page) return;
    const categories = state.categories.map(item => item.name);
    const rows = state.tasks.filter(task => {
      const ownerMatch = taskOwnerFilter === '全部' || task.owner === taskOwnerFilter;
      const status = task.completed ? '已完成' : '未完成';
      return ownerMatch && (taskStatusFilter === '全部' || taskStatusFilter === status);
    });
    page.innerHTML = `
      ${pageHead('HQ Task Board', '全新2.0任务记录', '旧Task Board保留在原Sheet但不会载入这里。团队成员可留空；类别可使用Playbook标准，也可随时新增。', 'tasks')}
      <div class="hq-form">
        ${field('Task内容', 'hq-task-name', 'text', '', 'wide')}
        <div class="hq-field"><label>负责人</label><select id="hq-task-owner">${OWNER_OPTIONS.map(owner => `<option>${owner}</option>`).join('')}</select></div>
        ${field('团队成员（可留空）', 'hq-task-member', 'text', '')}
        <div class="hq-field"><label>类别</label><input id="hq-task-category" list="hq-task-categories" placeholder="选择或输入新类别"><datalist id="hq-task-categories">${categories.map(category => `<option value="${esc(category)}"></option>`).join('')}</datalist></div>
        ${field('Deadline', 'hq-task-deadline', 'date', '')}
        <button class="hq-btn" onclick="hqAddTask()">添加Task</button>
      </div>
      <div class="hq-callout" style="margin-bottom:10px"><b>Playbook Task Rate 类别：</b>${categories.map(esc).join(' · ')}</div>
      <div class="hq-filter-row">
        <span class="hq-note">负责人：</span>
        ${['全部'].concat(OWNER_OPTIONS).map(owner => `<button class="hq-filter ${owner === taskOwnerFilter ? 'on' : ''}" onclick="hqSetTaskOwner('${owner}')">${owner}</button>`).join('')}
        <span class="hq-note">状态：</span>
        ${['全部', '未完成', '已完成'].map(status => `<button class="hq-filter ${status === taskStatusFilter ? 'on' : ''}" onclick="hqSetTaskStatus('${status}')">${status}</button>`).join('')}
      </div>
      <div class="hq-table-wrap"><table class="hq-table">
        <thead><tr><th>完成</th><th>Task内容</th><th>负责人</th><th>团队成员</th><th>类别</th><th>Deadline</th><th>完成日期</th><th></th></tr></thead>
        <tbody>${rows.length ? rows.map(task => {
          const index = state.tasks.indexOf(task);
          return `<tr>
            <td><input type="checkbox" ${task.completed ? 'checked' : ''} onchange="hqToggleTask(${index},this.checked)"></td>
            <td><input value="${esc(task.task)}" onchange="hqUpdateTask(${index},'task',this.value)"></td>
            <td><select onchange="hqUpdateTask(${index},'owner',this.value)">${OWNER_OPTIONS.map(owner => `<option ${owner === task.owner ? 'selected' : ''}>${owner}</option>`).join('')}</select></td>
            <td><input value="${esc(task.team_member || '')}" onchange="hqUpdateTask(${index},'team_member',this.value)"></td>
            <td><input list="hq-task-categories" value="${esc(task.category || '')}" onchange="hqUpdateTask(${index},'category',this.value)"></td>
            <td><input type="date" value="${esc(task.deadline || '')}" onchange="hqUpdateTask(${index},'deadline',this.value)"></td>
            <td><input type="date" value="${esc(task.completed_date || '')}" onchange="hqUpdateTask(${index},'completed_date',this.value)"></td>
            <td><button class="hq-btn danger" onclick="hqDeleteTask(${index})">删除</button></td>
          </tr>`;
        }).join('') : `<tr><td colspan="8"><div class="hq-empty">新的 HQ Task Board 目前是空的。上方添加第一项2.0任务。</div></td></tr>`}</tbody>
      </table></div>`;
  }

  function hqAddTask() {
    const task = document.getElementById('hq-task-name').value.trim();
    if (!task) return;
    const category = document.getElementById('hq-task-category').value.trim();
    state.tasks.push({
      id: 'task-' + Date.now(),
      task,
      owner: document.getElementById('hq-task-owner').value,
      team_member: document.getElementById('hq-task-member').value.trim(),
      category,
      deadline: document.getElementById('hq-task-deadline').value,
      completed_date: '',
      completed: false,
      last_updated: isoNow()
    });
    if (category && !state.categories.some(item => item.name === category)) {
      state.categories.push({ id: 'cat-' + Date.now(), name: category });
      persist('categories');
    }
    persist('tasks');
    renderHQTasks();
    renderHQOverview();
  }

  function hqSetTaskOwner(owner) { taskOwnerFilter = owner; renderHQTasks(); }
  function hqSetTaskStatus(status) { taskStatusFilter = status; renderHQTasks(); }

  function hqUpdateTask(index, key, value) {
    if (!state.tasks[index]) return;
    state.tasks[index][key] = value;
    state.tasks[index].last_updated = isoNow();
    if (key === 'category' && value && !state.categories.some(item => item.name === value)) {
      state.categories.push({ id: 'cat-' + Date.now(), name: value });
      persist('categories');
    }
    persist('tasks');
    renderHQTasks();
  }

  function hqToggleTask(index, checked) {
    if (!state.tasks[index]) return;
    state.tasks[index].completed = checked;
    state.tasks[index].completed_date = checked ? TODAY : '';
    state.tasks[index].last_updated = isoNow();
    persist('tasks');
    renderHQTasks();
    renderHQOverview();
  }

  function hqDeleteTask(index) {
    state.tasks.splice(index, 1);
    persist('tasks');
    renderHQTasks();
    renderHQOverview();
  }

  function renderHQCalendar() {
    const page = document.getElementById('p-calendar');
    if (!page) return;
    const events = [...state.events].sort((a, b) => a.date.localeCompare(b.date));
    page.innerHTML = `
      ${pageHead('Calendar', 'Upcoming Delivery & Events', 'Phase 1使用手动/Sheet活动资料。Google Calendar接口已预留，但不会显示假同步。', 'events')}
      <div class="hq-form">
        ${field('活动名称', 'hq-event-name', 'text', '', 'wide')}
        ${field('日期', 'hq-event-date', 'date', '')}
        ${field('时间', 'hq-event-time', 'text', '')}
        <div class="hq-field"><label>类型</label><select id="hq-event-type"><option>Product Class</option><option>Workshop</option><option>Free Talk</option><option>Venue Booking</option><option>Deadline</option><option>Reminder</option><option>Planning</option></select></div>
        <div class="hq-field"><label>形式</label><select id="hq-event-mode"><option>Online</option><option>Physical</option><option>TBD</option></select></div>
        ${field('地点 / Link', 'hq-event-location', 'text', '')}
        <button class="hq-btn" onclick="hqAddEvent()">添加活动</button>
      </div>
      <div class="hq-panel" style="margin-bottom:12px">
        <div class="hq-panel-title"><span>活动规划规则</span><span class="hq-status blue">不制造假日期</span></div>
        <div class="hq-note">Physical Workshop 从13/8开始，目标每月3场，安排在周四或周六下午；未确定的具体场次不会自动生成。Bootcamp Free Talk 从19/8起，每期计划2-3个周二，待时间确定后再加入日历。</div>
      </div>
      <div class="hq-event-list">${events.map((event, index) => `
        <div class="hq-event">
          <div class="hq-event-date">${fmtDate(event.date)}<div class="hq-event-meta">${esc(event.time)}</div></div>
          <div><div class="hq-event-name">${esc(event.name)}</div><div class="hq-event-meta">${esc(event.type)} · ${esc(event.mode)} · ${esc(event.location)} · ${esc(event.owner || 'Founder')}</div></div>
          <button class="hq-btn danger" onclick="hqDeleteEvent(${index})">删除</button>
        </div>`).join('')}</div>`;
  }

  function hqAddEvent() {
    const name = document.getElementById('hq-event-name').value.trim();
    const date = document.getElementById('hq-event-date').value;
    if (!name || !date) return;
    state.events.push({
      id: 'event-' + Date.now(),
      name,
      date,
      time: document.getElementById('hq-event-time').value,
      type: document.getElementById('hq-event-type').value,
      mode: document.getElementById('hq-event-mode').value,
      location: document.getElementById('hq-event-location').value,
      owner: 'Founder'
    });
    persist('events');
    renderHQCalendar();
    renderHQOverview();
  }

  function hqDeleteEvent(index) {
    state.events.splice(index, 1);
    persist('events');
    renderHQCalendar();
    renderHQOverview();
  }

  function renderSuccessPage() {
    const page = ensurePage('hq-success');
    const promises = [
      '看清真实财务全貌',
      '计算真实时薪',
      '找出消费模式',
      '算出自由数字',
      '建立可持续的财富系统',
      'Money-Life Templates'
    ];
    page.innerHTML = `
      ${pageHead('Customer Success', 'Product Excellence & Delivery', '不是为了记录投诉，而是确保承诺真的被交付，并把每一期变得更好。', 'success')}
      <div class="hq-grid-2">
        <div class="hq-panel">
          <div class="hq-panel-title"><span>第一期承诺成果</span><span class="hq-status green">6周</span></div>
          <div class="hq-promise-list">${promises.map(item => `<div class="hq-promise">✓ ${esc(item)}</div>`).join('')}</div>
        </div>
        <div class="hq-panel">
          <div class="hq-panel-title"><span>准备进度与风险要记录什么？</span></div>
          <div class="hq-note"><b>准备进度：</b>课程内容完成度、讲师/负责人、Zoom或课室、教材Template、付款名单、提醒讯息、录影和回播权限是否测试。</div>
          <div class="hq-note" style="margin-top:9px"><b>风险：</b>待付款、出席不足、内容未完成、链接未测试、负责人不明确、日期冲突，或客户承诺无法在期限内交付。</div>
        </div>
      </div>
      ${state.success.map((row, index) => `
        <div class="hq-panel" style="margin-bottom:12px">
          <div class="hq-panel-title"><span>${esc(row.product)}</span>${statusBadge(row.preparation)}</div>
          <div class="hq-grid-2">
            <div><div class="hq-page-kicker">Promised Result</div><div class="hq-note">${esc(row.promised_result)}</div></div>
            <div><div class="hq-page-kicker">Next Delivery</div><div style="font-size:13px;font-weight:800">${esc(row.next_delivery)}</div></div>
          </div>
          <div class="hq-form" style="margin-bottom:0">
            ${field('准备进度', 'hq-success-prep-' + index, 'text', row.preparation)}
            ${field('风险', 'hq-success-risk-' + index, 'text', row.risk, 'wide')}
            ${field('客户反馈', 'hq-success-feedback-' + index, 'text', row.feedback || '')}
            ${field('观察到的问题', 'hq-success-problem-' + index, 'text', row.problem || '')}
            ${field('改善行动', 'hq-success-improve-' + index, 'text', row.improvement || '', 'wide')}
            <button class="hq-btn" onclick="hqSaveSuccess(${index})">保存</button>
          </div>
        </div>`).join('')}`;
  }

  function hqSaveSuccess(index) {
    const row = state.success[index];
    if (!row) return;
    row.preparation = document.getElementById('hq-success-prep-' + index).value;
    row.risk = document.getElementById('hq-success-risk-' + index).value;
    row.feedback = document.getElementById('hq-success-feedback-' + index).value;
    row.problem = document.getElementById('hq-success-problem-' + index).value;
    row.improvement = document.getElementById('hq-success-improve-' + index).value;
    row.last_updated = isoNow();
    persist('success');
    renderSuccessPage();
  }

  function renderOperationsPage() {
    const page = ensurePage('hq-operations');
    page.innerHTML = `
      ${pageHead('Operations', 'Links, Delivery & Office', '日常操作入口、收据、课程回播、Office场地与交付任务集中在这里。', 'operations')}
      <div class="hq-grid-2">
        <div class="hq-panel">
          <div class="hq-panel-title"><span>Quick Access</span></div>
          <div class="hq-links">${state.links.map(linkCard).join('')}</div>
        </div>
        <div class="hq-panel">
          <div class="hq-panel-title"><span>DLIG Office 场地</span><span class="hq-status green">自有场地</span></div>
          <div style="font-size:18px;font-weight:800">约624 sqft · 约50人</div>
          <div class="hq-note" style="margin-top:8px">Projector、音响、麦克风×2、白板、椅子×50、双人桌×6。活动前后各45分钟布置/收拾。</div>
          <div class="hq-callout" style="margin-top:12px">Public：最低RM150（最多2.5小时），之后RM60/小时；8小时RM350；12小时RM450。正餐另加RM50清洁费。</div>
        </div>
      </div>
      <div class="hq-grid-2">
        <div class="hq-panel">
          <div class="hq-panel-title"><span>Task Fee / Standard Rate</span><span class="hq-status blue">Strategy Playbook</span></div>
          <div class="hq-table-wrap"><table class="hq-table" style="min-width:560px">
            <thead><tr><th>Task</th><th>Standard Rate</th></tr></thead>
            <tbody>
              <tr><td>AI Poster + 文案</td><td>RM75 / task</td></tr>
              <tr><td>AI Landing Page + Form + Payment + Auto Message</td><td>RM300 / task</td></tr>
              <tr><td>Social Media Sales AI Content（15 posts + 广告设置与监管）</td><td>RM1,200 / month</td></tr>
              <tr><td>Customer Service - WA、FB、小红书</td><td>RM700 / month</td></tr>
              <tr><td>Admin & Account</td><td>RM700 / month</td></tr>
              <tr><td>Event Space Rental Commission</td><td>20%</td></tr>
              <tr><td>Sales Commission / Event Crew</td><td>视活动而定</td></tr>
            </tbody>
          </table></div>
        </div>
        <div class="hq-panel">
          <div class="hq-panel-title"><span>Payment Status</span></div>
          <div class="hq-priority"><div class="hq-priority-num">P</div><div><div class="hq-priority-text">Paid</div><div class="hq-priority-sub">公司有预算，完成后付款</div></div></div>
          <div class="hq-priority"><div class="hq-priority-num">D</div><div><div class="hq-priority-text">Deferred</div><div class="hq-priority-sub">记录金额，项目收到钱后优先支付</div></div></div>
          <div class="hq-priority"><div class="hq-priority-num">C</div><div><div class="hq-priority-text">Contribution</div><div class="hq-priority-sub">事前同意义务贡献，不记录为公司欠款</div></div></div>
          <div class="hq-callout" style="margin-top:10px"><b>每个任务开始前先确认Payment Status。</b>Startup Contribution Rate最低以10%谈；收入达到内部门槛后才考虑Top-up。</div>
        </div>
      </div>
      ${editableActionTable(state.operations, 'operations')}`;
  }

  function linkCard(link) {
    return `<a class="hq-link-card" href="${esc(link.url)}" target="_blank" rel="noopener noreferrer">
      <span class="hq-link-icon">${esc(link.icon)}</span>
      <span><span class="hq-link-name">${esc(link.name)}</span><span class="hq-link-note">${esc(link.note)}</span></span>
    </a>`;
  }

  function renderIdeasPage() {
    const page = ensurePage('hq-ideas');
    page.innerHTML = `
      ${pageHead('Idea Parking Lot', '未来价值，不占用今日专注', '只有状态改为Approved的Idea才可以进入Daily Command Centre；Parking或待评估不会成为活跃优先事项。', 'ideas')}
      <div class="hq-form">
        ${field('Idea', 'hq-idea-name', 'text', '', 'wide')}
        ${field('Reason', 'hq-idea-reason', 'text', '')}
        ${field('Potential Value', 'hq-idea-value', 'text', '')}
        ${field('Revisit Date', 'hq-idea-date', 'date', '')}
        <button class="hq-btn" onclick="hqAddIdea()">Park Idea</button>
      </div>
      <div class="hq-table-wrap"><table class="hq-table">
        <thead><tr><th>Idea</th><th>Reason</th><th>Potential Value</th><th>Revisit</th><th>Status</th><th></th></tr></thead>
        <tbody>${state.ideas.map((idea, index) => `<tr>
          <td><b>${esc(idea.idea)}</b></td>
          <td>${esc(idea.reason)}</td>
          <td>${esc(idea.potential_value)}</td>
          <td><input type="date" value="${esc(idea.revisit_date || '')}" onchange="hqUpdateIdea(${index},'revisit_date',this.value)"></td>
          <td><select onchange="hqUpdateIdea(${index},'status',this.value)">${['Parking', '待评估', 'Approved', 'Rejected', 'Done'].map(status => `<option ${status === idea.status ? 'selected' : ''}>${status}</option>`).join('')}</select></td>
          <td><button class="hq-btn danger" onclick="hqDeleteIdea(${index})">删除</button></td>
        </tr>`).join('')}</tbody>
      </table></div>`;
  }

  function hqAddIdea() {
    const idea = document.getElementById('hq-idea-name').value.trim();
    if (!idea) return;
    state.ideas.push({
      id: 'idea-' + Date.now(),
      idea,
      reason: document.getElementById('hq-idea-reason').value,
      potential_value: document.getElementById('hq-idea-value').value,
      revisit_date: document.getElementById('hq-idea-date').value,
      status: 'Parking'
    });
    persist('ideas');
    renderIdeasPage();
  }

  function hqUpdateIdea(index, key, value) {
    if (!state.ideas[index]) return;
    state.ideas[index][key] = value;
    persist('ideas');
    renderIdeasPage();
  }

  function hqDeleteIdea(index) {
    state.ideas.splice(index, 1);
    persist('ideas');
    renderIdeasPage();
  }

  function renderHQFiles() {
    const page = document.getElementById('p-files');
    if (!page) return;
    page.innerHTML = `
      ${pageHead('File Index', '2.0 Quick Links + 1.0 Archive', '旧产品资料仍然可以查看，但不会占用2.0主导航或今日优先事项。', 'links')}
      <div class="hq-panel" style="margin-bottom:12px">
        <div class="hq-panel-title"><span>DLIG 2.0 工作入口</span></div>
        <div class="hq-links">${state.links.map(linkCard).join('')}</div>
      </div>
      <div class="hq-panel">
        <div class="hq-panel-title"><span>📦 2025 DLIG 1.0 产品记录</span><span class="hq-status gray">历史资料</span></div>
        <div class="hq-archive-grid">
          ${archiveCard('sales', 'Sales & Marketing', '旧销售记录与Marketing资料')}
          ${archiveCard('inventory', 'Inventory', '1.0产品物料与库存')}
          ${archiveCard('ylyd', 'YLYD System', '旧YLYD System资料')}
          ${archiveCard('bs', 'YLYD 会员', '会员流程与轮值')}
          ${archiveCard('xindong', '心动觉察', '旧产品交付与流程')}
          ${archiveCard('products', '产品与价格', '1.0产品价格记录')}
          ${archiveCard('admin', 'Admin & Account', '旧行政、费用与付款记录')}
          ${archiveCard('policy', 'Company Policy', '公司政策历史记录')}
          ${archiveCard('meeting', '会议与记录', '旧会议流程与记录')}
          ${archiveCard('decisions', '待决策记录', '1.0时期讨论与决定')}
        </div>
        <div class="hq-note" style="margin-top:12px">GDC Marketing Job 已从2.0界面移除，不再出现在导航或文件索引。</div>
      </div>`;
  }

  function archiveCard(id, title, note) {
    return `<div class="hq-archive-card" onclick="hqOpenLegacy('${id}','${esc(title)}')">
      <span><span class="hq-archive-label">${esc(title)}</span><span class="hq-archive-note">${esc(note)}</span></span><span>›</span>
    </div>`;
  }

  function renderById(id) {
    if (id === 'overview') renderHQOverview();
    if (id === 'hq-revenue') renderRevenuePage();
    if (id === 'hq-pipeline') renderPipelinePage();
    if (id === 'hq-daily') renderDailyPage();
    if (id === 'tasks') renderHQTasks();
    if (id === 'calendar') renderHQCalendar();
    if (id === 'hq-success') renderSuccessPage();
    if (id === 'hq-operations') renderOperationsPage();
    if (id === 'hq-ideas') renderIdeasPage();
    if (id === 'files') renderHQFiles();
  }

  function renderCurrent() {
    const current = (location.hash || '#overview').slice(1);
    renderById(current.startsWith('legacy-') ? 'files' : current);
  }

  function renderAllHQ() {
    renderHQOverview();
    renderRevenuePage();
    renderPipelinePage();
    renderDailyPage();
    renderHQTasks();
    renderHQCalendar();
    renderSuccessPage();
    renderOperationsPage();
    renderIdeasPage();
    renderHQFiles();
  }

  function registerGlobals() {
    window.hqGo = hqGo;
    window.hqOpenLegacy = openLegacy;
    window.hqSaveRevenue = hqSaveRevenue;
    window.hqUpdateProduct = hqUpdateProduct;
    window.hqUpdateMonthly = hqUpdateMonthly;
    window.hqSetPipelineFilter = hqSetPipelineFilter;
    window.hqAddPipeline = hqAddPipeline;
    window.hqUpdatePipeline = hqUpdatePipeline;
    window.hqDeletePipeline = hqDeletePipeline;
    window.hqSavePriority = hqSavePriority;
    window.hqAddDaily = hqAddDaily;
    window.hqUpdateAction = hqUpdateAction;
    window.hqDeleteAction = hqDeleteAction;
    window.hqAddTask = hqAddTask;
    window.hqSetTaskOwner = hqSetTaskOwner;
    window.hqSetTaskStatus = hqSetTaskStatus;
    window.hqUpdateTask = hqUpdateTask;
    window.hqToggleTask = hqToggleTask;
    window.hqDeleteTask = hqDeleteTask;
    window.hqAddEvent = hqAddEvent;
    window.hqDeleteEvent = hqDeleteEvent;
    window.hqSaveSuccess = hqSaveSuccess;
    window.hqAddIdea = hqAddIdea;
    window.hqUpdateIdea = hqUpdateIdea;
    window.hqDeleteIdea = hqDeleteIdea;
  }

  function init() {
    document.body.classList.add('hq2');
    setNavigation();
    setTopbar();
    [
      'hq-revenue', 'hq-pipeline', 'hq-daily', 'hq-success',
      'hq-operations', 'hq-ideas'
    ].forEach(ensurePage);
    registerGlobals();

    // Replace only the visible 2.0 renderers. Legacy page renderers remain intact.
    renderOverview = renderHQOverview;
    renderTasks = renderHQTasks;
    renderCalendar = renderHQCalendar;
    renderFiles = renderHQFiles;
    go = hqGo;

    renderAllHQ();
    const hash = location.hash.slice(1);
    const valid = ['overview', 'hq-revenue', 'hq-pipeline', 'hq-daily', 'tasks', 'calendar', 'hq-success', 'hq-operations', 'hq-ideas', 'files'];
    const legacyMatch = hash.match(/^legacy-(products|sales|admin|inventory|ylyd|bs|xindong|policy|meeting|decisions)$/);
    if (legacyMatch) {
      const legacyTitles = {
        products: '产品 & 价格',
        sales: 'Sales & Marketing',
        admin: 'Admin & Account',
        inventory: 'Inventory',
        ylyd: 'YLYD System',
        bs: 'YLYD 会员',
        xindong: '心动觉察',
        policy: 'Company Policy',
        meeting: '会议与记录',
        decisions: '待决策记录'
      };
      openLegacy(legacyMatch[1], legacyTitles[legacyMatch[1]]);
    } else {
      hqGo(valid.includes(hash) ? hash : 'overview', true);
    }
    loadHQData();
  }

  init();
})();
