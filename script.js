// Signup Department: show custom input when 'Others' selected
(function () {
  const sDeptSel = document.getElementById('signupDepartment');
  const sDeptOther = document.getElementById('signupDepartmentOther');
  if (sDeptSel && sDeptOther) {
    const toggle = () => {
      const show = sDeptSel.value === 'Others';
      sDeptOther.style.display = show ? 'block' : 'none';
      sDeptOther.required = show;
      if (!show) sDeptOther.value = '';
    };
    sDeptSel.addEventListener('change', toggle);
    // initialize
    try { toggle(); } catch (_) { }
  }
})();
/*
 * Construction Project Monitoring Dashboard Script
 * Handles CRUD operations, filtering, exporting, and rendering charts
 * Data persistence uses localStorage for demo purposes.
 */

(() => {
  const constants = window.AppConstants;
  if (!constants) {
    throw new Error('AppConstants missing; ensure js/constants.js loads before script.js');
  }
  const {
    LS_KEY,
    PROJECTS_COL,
    DEEPWELLS_COL,
    REFORESTATION_COL,
    SERVICE_UPDATES_COL,
    PERSONAL_TASKS_COL,
    PRESENTATIONS_COL,
    ADMIN_EMAIL,
    ADMIN_EMAIL_LOWER,
    ADMIN_FULL_NAME,
    ADMIN_DESIGNATION,
    ADMIN_DEPARTMENT,
    PERSONAL_TASKS_PER_PAGE,
    PROJECTS_PER_PAGE,
    DEEPWELLS_PER_PAGE,
    SERVICE_UPDATES_PER_PAGE,
    MWCI_PLANTS,
    MWSI_PLANTS,
  } = constants;

  const firebaseSvc = window.AppFirebase;
  if (!firebaseSvc) {
    throw new Error('AppFirebase missing; ensure js/services/firebase.js loads before script.js');
  }
  const {
    db,
    auth,
    FieldValue,
    Timestamp,
    serverTimestamp,
  } = firebaseSvc;

  const projectService = window.ProjectService;
  if (!projectService) {
    throw new Error('ProjectService missing; ensure js/services/projects.js loads before script.js');
  }
  const deepwellService = window.DeepwellService;
  const serviceUpdateService = window.ServiceUpdateService;
  const personalTaskService = window.PersonalTaskService;
  const presentationService = window.PresentationService;
  const opcrService = window.OPCRService;
  const ipcrService = window.IPCRService;
  const calendarService = window.CalendarService;
  const userService = window.UserService;
  const messagingService = window.MessagingService;
  if (!deepwellService || !serviceUpdateService || !personalTaskService || !presentationService || !opcrService || !ipcrService) {
    throw new Error('Collection services missing; ensure js/services/collections.js loads before script.js');
  }
  if (!userService) throw new Error('UserService missing; ensure js/services/users.js loads before script.js');
  if (!messagingService) throw new Error('MessagingService missing; ensure js/services/messaging.js loads before script.js');
  const emailService = window.EmailService;
  if (!emailService) throw new Error('EmailService missing; ensure js/services/collections.js loads before script.js');

  const utils = window.AppUtils;
  if (!utils) {
    throw new Error('AppUtils missing; ensure js/utils.js loads before script.js');
  }
  const {
    fmtNum,
    escapeHtml,
    escapeAttr,
    normalizeYmd,
    safeDateCompare,
    canonText,
    todayYmd,
    jan1OfYear,
    dec31OfYear,
    debounce,
    daysInMonthStr,
  } = utils;
  // ---- Personal Task Monitoring Helpers ----
  // Global loading spinner helper
  window.showLoading = (show) => {
    const spinner = document.getElementById('loadingSpinner');
    if (!spinner) return;
    if (show) spinner.classList.add('show');
    else spinner.classList.remove('show');
  };

  function showPersonalTasksSection() {
    // Fade out current content
    const mainContent = document.getElementById('mainContent');
    mainContent.classList.add('fade-out');

    setTimeout(() => {
      try { elements.dashboardSection.style.display = 'none'; } catch (_) { }
      try { elements.projectsSection.style.display = 'none'; } catch (_) { }
      try { elements.deepwellsSection.style.display = 'none'; } catch (_) { }
      try { elements.reforestationSection.style.display = 'none'; } catch (_) { }
      try { elements.serviceUpdateSection.style.display = 'none'; } catch (_) { }
      try { elements.presentationsSection.style.display = 'none'; } catch (_) { }
      try { elements.calendarSection.style.display = 'none'; } catch (_) { }
      try { elements.opcrSection.style.display = 'none'; } catch (_) { }
      try { elements.opcrFormSection.style.display = 'none'; } catch (_) { }
      try { elements.ipcrSection.style.display = 'none'; } catch (_) { }
      try { document.getElementById('documentRegistrySection').style.display = 'none'; } catch (_) { }

      // Update tabs
      if (elements.dashboardTab) elements.dashboardTab.classList.remove('active');
      if (elements.projectsTab) elements.projectsTab.classList.remove('active');
      if (elements.deepwellsTab) elements.deepwellsTab.classList.remove('active');
      if (elements.reforestationTab) elements.reforestationTab.classList.remove('active');
      if (elements.serviceUpdateTab) elements.serviceUpdateTab.classList.remove('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.remove('active');
      if (elements.calendarTab) elements.calendarTab.classList.remove('active');
      if (elements.opcrTab) elements.opcrTab.classList.remove('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.remove('active');

      if (elements.personalTasksSection) {
        elements.personalTasksSection.style.display = 'block';
        elements.personalTasksSection.classList.add('fade-element');
      }
      renderPersonalTasks();

      // Fade in
      mainContent.classList.remove('fade-out');
    }, 200); // Match CSS transition speed roughly
  }

  function getDamAvailableYears() {
    try {
      const years = new Set();
      const list = (serviceUpdates || []);
      for (const s of list) {
        const d = (s?.date || s?.timestamp || s?.createdAt);
        const y = (normalizeYmd ? normalizeYmd(d) : String(d || '')).slice(0, 4);
        if (y && /^\d{4}$/.test(y)) years.add(parseInt(y, 10));
      }
      return Array.from(years).sort((a, b) => b - a);
    } catch (_) { return []; }
  }

  function populateDamYearSelect() {
    const sel = document.getElementById('damYearSelect');
    if (!sel) return;
    const years = getDamAvailableYears();
    const signature = JSON.stringify(years);
    if (dashDamYearSelectSignature === signature) return;
    dashDamYearSelectSignature = signature;

    const current = new Date().getFullYear();
    sel.innerHTML = '';
    years.forEach(y => {
      const opt = document.createElement('option');
      opt.value = y;
      opt.textContent = y;
      sel.appendChild(opt);
    });
    const prefer = years.includes(current) ? current : (years[0] || current);
    if (prefer) sel.value = String(prefer);
  }

  function renderDamElevationsChartDash() {
    const canvas = elements.damElevationsChartDashCanvas;
    const emptyEl = elements.dashDamChartEmpty;
    if (!canvas) return;
    const toNum = (v) => {
      if (v === null || v === undefined || v === '') return NaN;
      const n = (typeof v === 'number') ? v : parseFloat(v.toString().replace(/,/g, ''));
      return Number.isFinite(n) ? n : NaN;
    };
    const normYmd = (v) => {
      try {
        if (typeof v?.toDate === 'function') { const d = v.toDate(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
        if (typeof v === 'object' && typeof v.seconds === 'number') { const d = new Date(v.seconds * 1000); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
        const s = normalizeYmd ? normalizeYmd(v) : String(v || '');
        return s;
      } catch (_) { return ''; }
    };
    const toMillis = (v) => {
      try {
        if (typeof v?.toDate === 'function') { const d = v.toDate(); return d.getTime(); }
        if (typeof v === 'object' && typeof v.seconds === 'number') { return v.seconds * 1000; }
        const d = new Date(v);
        const t = d.getTime();
        return Number.isFinite(t) ? t : NaN;
      } catch (_) { return NaN; }
    };
    const list = (serviceUpdates || []).filter(s => (toNum(s.angat) || toNum(s.ipo) || toNum(s.laMesa)) && (s.date || s.timestamp || s.createdAt));
    if (!list.length) { if (emptyEl) emptyEl.style.display = 'block'; canvas.style.display = 'none'; if (dashDamChart) { dashDamChart.destroy(); dashDamChart = null; } return; }
    let latestYear = (() => {
      const sel = document.getElementById('damYearSelect');
      const v = sel && sel.value ? parseInt(sel.value, 10) : NaN;
      if (Number.isFinite(v)) return v;
      const yrs = getDamAvailableYears();
      const cur = new Date().getFullYear();
      if (yrs.includes(cur)) return cur;
      return yrs[0] || cur;
    })();
    const years = list.map(s => { const d = normYmd(s.date || s.timestamp || s.createdAt); return d ? parseInt(d.slice(0, 4), 10) : NaN; }).filter(n => Number.isFinite(n));
    if (!years.includes(latestYear) && years.length) { latestYear = Math.max(...years); }
    const sortedSource = list.slice().sort((a, b) => {
      const ta = toMillis(a.date || a.timestamp || a.createdAt);
      const tb = toMillis(b.date || b.timestamp || b.createdAt);
      return (Number.isFinite(ta) ? ta : 0) - (Number.isFinite(tb) ? tb : 0);
    });
    const dailyMap = new Map();
    sortedSource.forEach(s => {
      const rawDate = s.date || s.timestamp || s.createdAt;
      const dstrFull = normYmd(rawDate);
      if (!dstrFull || !dstrFull.startsWith(String(latestYear) + '-')) return;
      const dstr = dstrFull.slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dstr)) return;
      const a = toNum(s.angat);
      const i = toNum(s.ipo);
      const l = toNum(s.laMesa);
      if (!Number.isFinite(a) && !Number.isFinite(i) && !Number.isFinite(l)) return;
      const existing = dailyMap.get(dstr) || { angat: null, ipo: null, laMesa: null };
      if (Number.isFinite(a)) existing.angat = a;
      if (Number.isFinite(i)) existing.ipo = i;
      if (Number.isFinite(l)) existing.laMesa = l;
      dailyMap.set(dstr, existing);
    });
    const dateLabels = Array.from(dailyMap.keys()).sort((a, b) => a.localeCompare(b));
    const angat = dateLabels.map(d => dailyMap.get(d)?.angat ?? null);
    const ipo = dateLabels.map(d => dailyMap.get(d)?.ipo ?? null);
    const laMesa = dateLabels.map(d => dailyMap.get(d)?.laMesa ?? null);

    // Create a data signature to avoid re-rendering if data is the same
    const signature = JSON.stringify(dateLabels) + JSON.stringify(angat) + JSON.stringify(ipo) + JSON.stringify(laMesa);
    if (dashDamChartSignature === signature && dashDamChart) return;

    const hasAny = [...angat, ...ipo, ...laMesa].some(v => Number.isFinite(v));
    if (!hasAny) { if (emptyEl) emptyEl.style.display = 'block'; canvas.style.display = 'none'; if (dashDamChart) { dashDamChart.destroy(); dashDamChart = null; } return; }
    if (emptyEl) emptyEl.style.display = 'none'; canvas.style.display = 'block';
    const isMobile = window.matchMedia('(max-width: 576px)').matches;
    const tickFontSize = isMobile ? 10 : 12;
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const formatMonth = (dstr) => {
      if (typeof dstr !== 'string' || dstr.length < 7) return dstr;
      const m = parseInt(dstr.slice(5, 7), 10) - 1;
      return (m >= 0 && m < 12) ? monthNames[m] : dstr;
    };
    const formatFullDate = (dstr) => {
      if (typeof dstr !== 'string') return '';
      const [yy, mm, dd] = dstr.split('-').map(part => parseInt(part, 10));
      if (!Number.isFinite(yy) || !Number.isFinite(mm) || !Number.isFinite(dd)) return dstr;
      const dt = new Date(yy, mm - 1, dd);
      if (!Number.isFinite(dt.getTime())) return dstr;
      return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    };
    let lastTickMonth = null;
    const data = {
      labels: dateLabels,
      datasets: [
        { label: 'Angat', data: angat, borderColor: '#0d6efd', backgroundColor: 'rgba(13,110,253,0.15)', tension: 0.25, fill: false, borderWidth: 2, pointRadius: 2 },
        { label: 'Ipo', data: ipo, borderColor: '#198754', backgroundColor: 'rgba(25,135,84,0.15)', tension: 0.25, fill: false, borderWidth: 2, pointRadius: 2 },
        { label: 'La Mesa', data: laMesa, borderColor: '#fd7e14', backgroundColor: 'rgba(253,126,20,0.15)', tension: 0.25, fill: false, borderWidth: 2, pointRadius: 2 },
      ]
    };
    const opts = {
      responsive: true,
      maintainAspectRatio: true,
      aspectRatio: isMobile ? 1.6 : 2.2,
      resizeDelay: 150,
      interaction: { mode: 'index', intersect: false },
      layout: { padding: { top: 4, right: 6, bottom: 2, left: 4 } },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: tickFontSize } } },
        tooltip: {
          callbacks: {
            title: (ctx) => {
              if (!ctx?.length) return '';
              return formatFullDate(ctx[0].label);
            },
            label: (ctx) => {
              if (ctx.parsed.y == null) return `${ctx.dataset.label}: No data`;
              return `${ctx.dataset.label}: ${ctx.parsed.y.toFixed(2)} masl`;
            }
          }
        },
        title: { display: true, text: `Daily dam elevations - ${latestYear}` }
      },
      scales: {
        x: {
          title: { display: true, text: 'Month' },
          ticks: {
            font: { size: tickFontSize },
            maxTicksLimit: 12,
            callback: function (value, index, ticks) {
              const label = (ticks?.[index]?.label ?? this.getLabelForValue?.(value) ?? value);
              const month = formatMonth(label);
              if (month === lastTickMonth) return '';
              lastTickMonth = month;
              return month;
            }
          },
          grid: { color: 'rgba(0,0,0,0.06)' }
        },
        y: { title: { display: true, text: 'masl' }, beginAtZero: false, grid: { color: 'rgba(0,0,0,0.06)' } }
      }
    };
    if (dashDamChart) { dashDamChart.data = data; dashDamChart.options = opts; dashDamChart.update(); }
    else { const ctx = canvas.getContext('2d'); dashDamChart = new Chart(ctx, { type: 'line', data, options: opts }); }
    dashDamChartSignature = signature;
  }

  function getTreatmentPlantAvailableYears(provider) {
    try {
      const prov = String(provider || '').toUpperCase();
      const years = new Set();
      const toNum = (v) => {
        if (v === null || v === undefined || v === '') return NaN;
        const n = (typeof v === 'number') ? v : parseFloat(String(v).replace(/,/g, ''));
        return Number.isFinite(n) ? n : NaN;
      };
      const normYmd = (v) => {
        try {
          if (typeof v?.toDate === 'function') {
            const d = v.toDate();
            return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
          }
          if (typeof v === 'object' && typeof v.seconds === 'number') {
            const d = new Date(v.seconds * 1000);
            return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
          }
          return normalizeYmd ? normalizeYmd(v) : String(v || '');
        } catch (_) { return ''; }
      };
      (serviceUpdates || []).forEach(s => {
        if (String(s?.provider || '').toUpperCase() !== prov) return;
        const hasPlantProduction = Array.isArray(s.plants) && s.plants.some(p => {
          const n = toNum(p?.production);
          return Number.isFinite(n) && n !== 0;
        });
        const total = toNum(s.production);
        const hasTotalProduction = Number.isFinite(total) && total !== 0;
        if (!hasPlantProduction && !hasTotalProduction) return;
        const d = normYmd(s.plantsDate || s.date || s.timestamp || s.createdAt);
        const y = (d || '').slice(0, 4);
        if (/^\d{4}$/.test(y)) years.add(parseInt(y, 10));
      });
      return Array.from(years).sort((a, b) => b - a);
    } catch (_) { return []; }
  }

  function populateTreatmentPlantYearSelect(provider) {
    const prov = String(provider || '').toUpperCase();
    const key = prov === 'MWSI' ? 'Mwsi' : 'Mwci';
    const sel = document.getElementById(`treatmentPlants${key}YearSelect`);
    if (!sel) return;
    const years = getTreatmentPlantAvailableYears(prov);
    const signature = JSON.stringify(years);
    if (dashTreatmentPlantYearSelectSignatures[prov] === signature) return;
    dashTreatmentPlantYearSelectSignatures[prov] = signature;

    const previous = parseInt(sel.value, 10);
    const serviceYear = parseInt(elements.serviceYearFilter?.value || '', 10);
    const current = new Date().getFullYear();
    sel.innerHTML = '';
    years.forEach(y => {
      const opt = document.createElement('option');
      opt.value = y;
      opt.textContent = y;
      sel.appendChild(opt);
    });
    const prefer = years.includes(previous)
      ? previous
      : (years.includes(serviceYear) ? serviceYear : (years.includes(current) ? current : years[0]));
    if (prefer) sel.value = String(prefer);
  }

  function getTreatmentPlantSelectedYear(provider) {
    const prov = String(provider || '').toUpperCase() === 'MWSI' ? 'MWSI' : 'MWCI';
    const key = prov === 'MWSI' ? 'Mwsi' : 'Mwci';
    const years = getTreatmentPlantAvailableYears(prov);
    const sel = document.getElementById(`treatmentPlants${key}YearSelect`);
    const v = sel && sel.value ? parseInt(sel.value, 10) : NaN;
    let selectedYear = Number.isFinite(v) ? v : NaN;
    if (!Number.isFinite(selectedYear)) {
      const serviceYear = parseInt(elements.serviceYearFilter?.value || '', 10);
      selectedYear = Number.isFinite(serviceYear) ? serviceYear : (years[0] || new Date().getFullYear());
    }
    if (years.length && !years.includes(selectedYear)) selectedYear = years[0];
    return selectedYear;
  }

  function buildTreatmentPlantProductionSeries(provider, selectedYear) {
    const prov = String(provider || '').toUpperCase() === 'MWSI' ? 'MWSI' : 'MWCI';
    const toNum = (v) => {
      if (v === null || v === undefined || v === '') return NaN;
      const n = (typeof v === 'number') ? v : parseFloat(String(v).replace(/,/g, ''));
      return Number.isFinite(n) ? n : NaN;
    };
    const normYmd = (v) => {
      try {
        if (typeof v?.toDate === 'function') {
          const d = v.toDate();
          return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        }
        if (typeof v === 'object' && typeof v.seconds === 'number') {
          const d = new Date(v.seconds * 1000);
          return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        }
        return normalizeYmd ? normalizeYmd(v) : String(v || '');
      } catch (_) { return ''; }
    };
    const toMillis = (v) => {
      try {
        if (typeof v?.toDate === 'function') return v.toDate().getTime();
        if (typeof v === 'object' && typeof v.seconds === 'number') return v.seconds * 1000;
        const t = new Date(v).getTime();
        return Number.isFinite(t) ? t : 0;
      } catch (_) { return 0; }
    };

    const plantNames = getProviderPlants(prov);
    const plantNameSet = new Set(plantNames.map(n => String(n).trim()));
    const dailyMap = new Map();
    const sortedSource = (serviceUpdates || [])
      .filter(s => String(s?.provider || '').toUpperCase() === prov)
      .sort((a, b) => toMillis(a.plantsDate || a.date || a.timestamp || a.createdAt) - toMillis(b.plantsDate || b.date || b.timestamp || b.createdAt));

    sortedSource.forEach(s => {
      const dstrFull = normYmd(s.plantsDate || s.date || s.timestamp || s.createdAt);
      if (!dstrFull || !dstrFull.startsWith(String(selectedYear) + '-')) return;
      const dstr = dstrFull.slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dstr)) return;
      const existing = dailyMap.get(dstr) || {};
      let hasPlantProduction = false;
      if (Array.isArray(s.plants)) {
        s.plants.forEach(p => {
          const name = String(p?.name || '').trim();
          if (!name || !plantNameSet.has(name)) return;
          const production = toNum(p?.production);
          if (!Number.isFinite(production)) return;
          existing[name] = production;
          if (production !== 0) hasPlantProduction = true;
        });
      }
      const totalProduction = toNum(s.production);
      if (!hasPlantProduction && Number.isFinite(totalProduction) && totalProduction !== 0) {
        existing['Total Production'] = totalProduction;
      }
      dailyMap.set(dstr, existing);
    });

    const dateLabels = Array.from(dailyMap.keys()).sort((a, b) => a.localeCompare(b));
    const hasLegacyTotal = dateLabels.some(d => Number.isFinite(dailyMap.get(d)?.['Total Production']));
    const seriesNames = hasLegacyTotal ? [...plantNames, 'Total Production'] : plantNames;
    const colors = ['#0d6efd', '#198754', '#fd7e14', '#6f42c1', '#20c997', '#dc3545', '#0dcaf0', '#ffc107', '#6c757d', '#6610f2'];
    const datasets = seriesNames.map((name, idx) => {
      const values = dateLabels.map(d => dailyMap.get(d)?.[name] ?? null);
      const hasValues = values.some(v => Number.isFinite(v) && v !== 0);
      if (!hasValues) return null;
      const color = colors[idx % colors.length];
      return {
        label: name,
        data: values,
        borderColor: color,
        backgroundColor: color + '26',
        tension: 0.25,
        fill: false,
        borderWidth: 2,
        pointRadius: 1.5,
        pointHoverRadius: 4,
        spanGaps: true,
      };
    }).filter(Boolean);
    const stats = datasets.map(ds => {
      const entries = ds.data
        .map((value, idx) => ({ value, date: dateLabels[idx] }))
        .filter(item => Number.isFinite(item.value));
      const total = entries.reduce((sum, item) => sum + item.value, 0);
      const avg = entries.length ? total / entries.length : 0;
      const threshold = avg * 0.8;
      const lowDays = entries.filter(item => item.value < threshold).length;
      const latest = entries.length ? entries[entries.length - 1] : { value: NaN, date: '' };
      const min = entries.reduce((acc, item) => item.value < acc.value ? item : acc, { value: +Infinity, date: '' });
      const latestRatio = avg > 0 && Number.isFinite(latest.value) ? latest.value / avg : 1;
      const status = !Number.isFinite(latest.value)
        ? 'No recent data'
        : (latest.value < threshold ? 'Below average' : (lowDays > 0 ? 'Has low days' : 'Stable'));
      return {
        plant: ds.label,
        latest: latest.value,
        latestDate: latest.date,
        avg,
        min: min.value === +Infinity ? NaN : min.value,
        minDate: min.date,
        lowDays,
        threshold,
        latestRatio,
        status,
      };
    }).sort((a, b) => {
      const ar = Number.isFinite(a.latestRatio) ? a.latestRatio : 99;
      const br = Number.isFinite(b.latestRatio) ? b.latestRatio : 99;
      if (ar !== br) return ar - br;
      return b.lowDays - a.lowDays;
    });
    return { provider: prov, year: selectedYear, labels: dateLabels, datasets, stats };
  }

  function formatTreatmentPlantFullDate(dstr) {
    if (typeof dstr !== 'string') return '';
    const [yy, mm, dd] = dstr.split('-').map(part => parseInt(part, 10));
    if (!Number.isFinite(yy) || !Number.isFinite(mm) || !Number.isFinite(dd)) return dstr;
    const dt = new Date(yy, mm - 1, dd);
    return Number.isFinite(dt.getTime()) ? dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : dstr;
  }

  function createTreatmentPlantChartOptions(provider, year, expanded) {
    const prov = String(provider || '').toUpperCase() === 'MWSI' ? 'MWSI' : 'MWCI';
    const isMobile = window.matchMedia('(max-width: 576px)').matches;
    const tickFontSize = isMobile ? 10 : 12;
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const formatMonth = (dstr) => {
      if (typeof dstr !== 'string' || dstr.length < 7) return dstr;
      const m = parseInt(dstr.slice(5, 7), 10) - 1;
      return (m >= 0 && m < 12) ? monthNames[m] : dstr;
    };
    let lastTickMonth = null;
    return {
      responsive: true,
      maintainAspectRatio: !expanded,
      aspectRatio: expanded ? undefined : (isMobile ? 1.6 : 2.2),
      resizeDelay: 150,
      interaction: { mode: 'index', intersect: false },
      layout: { padding: { top: 4, right: 6, bottom: 2, left: 4 } },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: tickFontSize } } },
        tooltip: {
          callbacks: {
            title: (ctx) => ctx?.length ? formatTreatmentPlantFullDate(ctx[0].label) : '',
            label: (ctx) => {
              if (ctx.parsed.y == null) return `${ctx.dataset.label}: No data`;
              return `${ctx.dataset.label}: ${fmtNum(ctx.parsed.y)} MLD`;
            }
          }
        },
        title: { display: true, text: `Daily treatment production - ${prov} ${year}` }
      },
      scales: {
        x: {
          title: { display: true, text: 'Month' },
          ticks: {
            font: { size: tickFontSize },
            maxTicksLimit: expanded ? 18 : 12,
            callback: function (value, index, ticks) {
              const label = (ticks?.[index]?.label ?? this.getLabelForValue?.(value) ?? value);
              const month = formatMonth(label);
              if (month === lastTickMonth) return '';
              lastTickMonth = month;
              return month;
            }
          },
          grid: { color: 'rgba(0,0,0,0.06)' }
        },
        y: { title: { display: true, text: 'MLD' }, beginAtZero: true, grid: { color: 'rgba(0,0,0,0.06)' } }
      }
    };
  }

  function renderTreatmentPlantProductionChartDash(provider) {
    const prov = String(provider || '').toUpperCase() === 'MWSI' ? 'MWSI' : 'MWCI';
    const key = prov === 'MWSI' ? 'Mwsi' : 'Mwci';
    const canvas = elements[`treatmentPlants${key}ChartDashCanvas`];
    const emptyEl = elements[`dashTreatmentPlants${key}Empty`];
    if (!canvas) return;
    const selectedYear = getTreatmentPlantSelectedYear(prov);
    const series = buildTreatmentPlantProductionSeries(prov, selectedYear);
    const signature = JSON.stringify([prov, selectedYear, series.labels, series.datasets.map(d => [d.label, d.data])]);
    if (dashTreatmentPlantChartSignatures[prov] === signature && dashTreatmentPlantCharts[prov]) return;

    if (!series.datasets.length) {
      if (emptyEl) emptyEl.style.display = 'block';
      canvas.style.display = 'none';
      if (dashTreatmentPlantCharts[prov]) {
        dashTreatmentPlantCharts[prov].destroy();
        dashTreatmentPlantCharts[prov] = null;
      }
      dashTreatmentPlantChartSignatures[prov] = signature;
      return;
    }
    if (emptyEl) emptyEl.style.display = 'none';
    canvas.style.display = 'block';

    const data = { labels: series.labels, datasets: series.datasets };
    const opts = createTreatmentPlantChartOptions(prov, selectedYear, false);
    if (dashTreatmentPlantCharts[prov]) {
      dashTreatmentPlantCharts[prov].data = data;
      dashTreatmentPlantCharts[prov].options = opts;
      dashTreatmentPlantCharts[prov].update();
    } else {
      dashTreatmentPlantCharts[prov] = new Chart(canvas.getContext('2d'), { type: 'line', data, options: opts });
    }
    dashTreatmentPlantChartSignatures[prov] = signature;
  }

  function renderTreatmentPlantExpandedChart(series) {
    const canvas = elements.treatmentPlantsExpandedChartCanvas;
    if (!canvas) return;
    const data = { labels: series.labels, datasets: series.datasets };
    const opts = createTreatmentPlantChartOptions(series.provider, series.year, true);
    if (treatmentPlantsExpandedChart) {
      treatmentPlantsExpandedChart.data = data;
      treatmentPlantsExpandedChart.options = opts;
      treatmentPlantsExpandedChart.resize();
      treatmentPlantsExpandedChart.update('none');
    } else {
      treatmentPlantsExpandedChart = new Chart(canvas.getContext('2d'), { type: 'line', data, options: opts });
    }
  }

  function getTreatmentPlantTotalProductionToDate(series) {
    return (series?.datasets || []).reduce((sum, ds) => {
      return sum + (ds.data || []).reduce((inner, value) => inner + (Number.isFinite(value) ? value : 0), 0);
    }, 0);
  }

  function getTreatmentPlantExpandedSummary(series) {
    const stats = series?.stats || [];
    return [
      { label: 'Plants with data', value: fmtNum(stats.length) },
      { label: 'Data days', value: fmtNum(series?.labels?.length || 0) },
      { label: 'Total production to date', value: `${fmtNum(+getTreatmentPlantTotalProductionToDate(series).toFixed(2))} ML` },
    ];
  }

  function renderTreatmentPlantExpandedAnalysis(series) {
    const stats = series.stats || [];
    const summary = getTreatmentPlantExpandedSummary(series);
    if (elements.treatmentPlantsExpandedSummary) {
      elements.treatmentPlantsExpandedSummary.innerHTML = summary.map(item => `
        <div class="col-6 col-lg-3">
          <div class="border rounded p-2 h-100">
            <div class="text-muted small">${escapeHtml(item.label)}</div>
            <div class="fw-bold">${escapeHtml(item.value)}</div>
          </div>
        </div>
      `).join('');
    }
    if (elements.treatmentPlantsExpandedTbody) {
      elements.treatmentPlantsExpandedTbody.innerHTML = stats.map(s => {
        const isBelow = Number.isFinite(s.latest) && s.latest < s.threshold;
        const hasLowDays = !isBelow && s.lowDays > 0;
        const badgeClass = isBelow ? 'bg-danger' : (hasLowDays ? 'bg-warning text-dark' : 'bg-success');
        const statusText = isBelow ? 'Below average' : (hasLowDays ? 'Has low days' : s.status);
        return `<tr>
          <td>${escapeHtml(s.plant)}</td>
          <td class="text-end">${Number.isFinite(s.avg) ? fmtNum(+s.avg.toFixed(2)) : '-'}</td>
          <td class="text-end">${fmtNum(s.lowDays || 0)}</td>
          <td><span class="badge ${badgeClass}">${escapeHtml(statusText)}</span></td>
        </tr>`;
      }).join('');
    }
    if (elements.treatmentPlantsExpandedSubtitle) {
      elements.treatmentPlantsExpandedSubtitle.textContent = '';
    }
  }

  function getTreatmentPlantExpandedChartImage() {
    try {
      const canvas = elements.treatmentPlantsExpandedChartCanvas;
      if (!canvas || !canvas.width || !canvas.height) return '';
      return canvas.toDataURL('image/png');
    } catch (_) { return ''; }
  }

  function buildTreatmentPlantReportHtml(series, chartImage) {
    const title = `Treatment Plants - ${series.provider} (${series.year})`;
    const summaryHtml = getTreatmentPlantExpandedSummary(series).map(item => `
      <div class="summary-card">
        <div class="summary-label">${escapeHtml(item.label)}</div>
        <div class="summary-value">${escapeHtml(item.value)}</div>
      </div>
    `).join('');
    const rowsHtml = (series.stats || []).map(s => {
      const isBelow = Number.isFinite(s.latest) && s.latest < s.threshold;
      const hasLowDays = !isBelow && s.lowDays > 0;
      const statusText = isBelow ? 'Below average' : (hasLowDays ? 'Has low days' : s.status);
      return `<tr>
        <td>${escapeHtml(s.plant)}</td>
        <td class="num">${Number.isFinite(s.avg) ? fmtNum(+s.avg.toFixed(2)) : '-'}</td>
        <td class="num">${fmtNum(s.lowDays || 0)}</td>
        <td>${escapeHtml(statusText)}</td>
      </tr>`;
    }).join('');
    return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Arial, sans-serif; color: #111827; margin: 28px; }
    h1 { font-size: 22px; margin: 0 0 14px; }
    h2 { font-size: 16px; margin: 22px 0 8px; }
    .summary { display: table; width: 100%; border-spacing: 8px; margin-bottom: 14px; }
    .summary-card { display: table-cell; border: 1px solid #d1d5db; padding: 10px; width: 33.33%; }
    .summary-label { color: #6b7280; font-size: 12px; margin-bottom: 4px; }
    .summary-value { font-weight: 700; font-size: 15px; }
    .chart { width: 6.5in; height: auto; display: block; margin: 8px 0 16px; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th { background: #0f4775; color: #fff; text-align: left; }
    th, td { border: 1px solid #d1d5db; padding: 6px 8px; }
    .num { text-align: right; }
    @media print { body { margin: 14mm; } }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  <div class="summary">${summaryHtml}</div>
  ${chartImage ? `<img class="chart" src="${chartImage}" width="624" style="width:6.5in;height:auto;display:block;" alt="Treatment plant production chart">` : ''}
  <h2>Low Production Review</h2>
  <table>
    <thead>
      <tr>
        <th>Plant</th>
        <th class="num">Average</th>
        <th class="num">Low Days</th>
        <th>Status</th>
      </tr>
    </thead>
    <tbody>${rowsHtml}</tbody>
  </table>
</body>
</html>`;
  }

  function printTreatmentPlantExpandedReport() {
    const series = treatmentPlantsExpandedSeries;
    if (!series) return;
    const reportHtml = buildTreatmentPlantReportHtml(series, getTreatmentPlantExpandedChartImage());
    const win = window.open('', '_blank');
    if (!win) {
      alert('Unable to open print window. Please allow pop-ups for this site.');
      return;
    }
    win.document.open();
    win.document.write(reportHtml);
    win.document.close();
    win.focus();
    setTimeout(() => {
      try { win.print(); } catch (_) { }
    }, 300);
  }

  function downloadTreatmentPlantExpandedWord() {
    const series = treatmentPlantsExpandedSeries;
    if (!series) return;
    const reportHtml = buildTreatmentPlantReportHtml(series, getTreatmentPlantExpandedChartImage());
    const blob = new Blob(['\ufeff' + reportHtml], { type: 'application/msword' });
    const provider = String(series.provider || 'Treatment').replace(/[^a-z0-9_-]+/gi, '_');
    const filename = `Treatment_Plants_${provider}_${series.year}.doc`;
    if (window.saveAs) {
      window.saveAs(blob, filename);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
  }

  function openTreatmentPlantExpanded(provider) {
    const prov = String(provider || '').toUpperCase() === 'MWSI' ? 'MWSI' : 'MWCI';
    const year = getTreatmentPlantSelectedYear(prov);
    const series = buildTreatmentPlantProductionSeries(prov, year);
    treatmentPlantsExpandedSeries = series;
    if (elements.treatmentPlantsExpandedTitle) {
      elements.treatmentPlantsExpandedTitle.textContent = `Treatment Plants - ${prov} (${year})`;
    }
    const hasData = series.datasets.length > 0;
    if (elements.treatmentPlantsExpandedEmpty) elements.treatmentPlantsExpandedEmpty.style.display = hasData ? 'none' : 'block';
    if (elements.treatmentPlantsExpandedChartCanvas) elements.treatmentPlantsExpandedChartCanvas.style.display = hasData ? 'block' : 'none';
    renderTreatmentPlantExpandedAnalysis(series);
    elements.treatmentPlantsExpandedModal.show();
    setTimeout(() => {
      if (hasData) renderTreatmentPlantExpandedChart(series);
    }, 180);
  }

  function wireTreatmentPlantExpandedCards() {
    const wireCard = (id, provider) => {
      const card = document.getElementById(id);
      if (!card || card.__treatmentPlantExpandedWired) return;
      const shouldIgnore = (target) => !!target.closest('select,label,option,.form-select');
      card.addEventListener('click', e => {
        if (shouldIgnore(e.target)) return;
        openTreatmentPlantExpanded(provider);
      });
      card.addEventListener('keydown', e => {
        if (shouldIgnore(e.target)) return;
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          openTreatmentPlantExpanded(provider);
        }
      });
      card.__treatmentPlantExpandedWired = true;
    };
    wireCard('treatmentPlantsMwciCard', 'MWCI');
    wireCard('treatmentPlantsMwsiCard', 'MWSI');
  }

  function wireTreatmentPlantExpandedExportActions() {
    if (elements.treatmentPlantsPrintBtn && !elements.treatmentPlantsPrintBtn.__wired) {
      elements.treatmentPlantsPrintBtn.addEventListener('click', printTreatmentPlantExpandedReport);
      elements.treatmentPlantsPrintBtn.__wired = true;
    }
    if (elements.treatmentPlantsWordBtn && !elements.treatmentPlantsWordBtn.__wired) {
      elements.treatmentPlantsWordBtn.addEventListener('click', downloadTreatmentPlantExpandedWord);
      elements.treatmentPlantsWordBtn.__wired = true;
    }
  }

  function normalizeDateStr(d) { return (d || '').slice(0, 10); }

  function personalTaskRowHtml(t) {
    const due = t.due || '';
    const today = new Date().toISOString().slice(0, 10);
    const isOverdue = due && t.status !== 'Done' && safeDateCompare(due, today) < 0;
    const prog = (t.status === 'Done') ? 100 : Math.max(0, Math.min(100, Number(t.progress || 0)));
    const badge = isOverdue ? '<span class="badge bg-danger ms-1">Overdue</span>' : '';
    const actions = `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="editPersonalTask('${t.id}')"><i class="fa fa-pencil"></i></button>` +
      `<button class="btn btn-sm btn-danger" title="Delete" onclick="deletePersonalTask('${t.id}')"><i class="fa fa-trash"></i></button>`;
    return `<tr data-id="${t.id}">`
      + `<td>${escapeHtml(t.title || '')}</td>`
      + `<td>${escapeHtml(t.category || '')}</td>`
      + `<td>${escapeHtml(t.priority || '')}</td>`
      + `<td>${window.AppUtils && window.AppUtils.formatDateUI ? window.AppUtils.formatDateUI(due) : (due || '')} ${badge}</td>`
      + `<td>${escapeHtml(t.status || '')}</td>`
      + `<td><div class="progress" style="height:8px;"><div class="progress-bar" role="progressbar" style="width:${prog}%;" aria-valuenow="${prog}" aria-valuemin="0" aria-valuemax="100"></div></div></td>`
      + `<td>${window.AppUtils && window.AppUtils.formatDateUI ? window.AppUtils.formatDateUI(normalizeDateStr(t.dateCompleted || '')) : normalizeDateStr(t.dateCompleted || '')}</td>`
      + `<td class="text-nowrap">${actions}</td>`
      + `</tr>`;
  }

  function personalTaskCardHtml(t) {
    const due = t.due || '';
    const today = new Date().toISOString().slice(0, 10);
    const isOverdue = due && t.status !== 'Done' && safeDateCompare(due, today) < 0;
    const prog = (t.status === 'Done') ? 100 : Math.max(0, Math.min(100, Number(t.progress || 0)));
    const overdueBadge = isOverdue ? '<span class="badge bg-danger ms-1">Overdue</span>' : '';
    
    // Priority formatting
    let priorityClass = 'bg-secondary-subtle text-secondary';
    const prio = (t.priority || '').toLowerCase();
    if (prio === 'high') {
      priorityClass = 'bg-warning-subtle text-warning-emphasis border border-warning-subtle';
    } else if (prio === 'urgent') {
      priorityClass = 'bg-danger-subtle text-danger border border-danger-subtle';
    } else if (prio === 'medium') {
      priorityClass = 'bg-primary-subtle text-primary border border-primary-subtle';
    } else if (prio === 'low') {
      priorityClass = 'bg-light text-dark border border-secondary-subtle';
    }

    // Status formatting
    let statusClass = 'bg-secondary-subtle text-secondary';
    const status = (t.status || '');
    if (status === 'Done') {
      statusClass = 'bg-success-subtle text-success border border-success-subtle';
    } else if (status === 'Ongoing') {
      statusClass = 'bg-info-subtle text-info-emphasis border border-info-subtle';
    } else if (status === 'Blocked') {
      statusClass = 'bg-danger-subtle text-danger border border-danger-subtle';
    } else if (status === 'Not Started') {
      statusClass = 'bg-light text-dark border border-secondary-subtle';
    }

    const fmtUi = window.AppUtils && window.AppUtils.formatDateUI ? window.AppUtils.formatDateUI : (x => x || '');
    const dueFormatted = fmtUi(due);
    const completedFormatted = t.dateCompleted ? fmtUi(normalizeDateStr(t.dateCompleted)) : '';

    const actions = `<button class="btn btn-sm btn-outline-primary me-2" title="Edit" onclick="event.stopPropagation(); editPersonalTask('${t.id}')"><i class="fa fa-pencil"></i> Edit</button>` +
      `<button class="btn btn-sm btn-outline-danger" title="Delete" onclick="event.stopPropagation(); deletePersonalTask('${t.id}')"><i class="fa fa-trash"></i> Delete</button>`;

    return `
      <div class="card mb-3 border border-light shadow-sm" data-id="${t.id}" style="cursor:pointer; border-radius: 12px; transition: transform 0.15s ease, box-shadow 0.15s ease;">
        <div class="card-body p-3">
          <div class="d-flex justify-content-between align-items-start mb-2 gap-1 flex-wrap">
            <span class="badge ${priorityClass} px-2 py-1 small rounded-pill">${escapeHtml(t.priority || 'Medium')} Priority</span>
            <span class="badge ${statusClass} px-2 py-1 small rounded-pill">${escapeHtml(status || 'Not Started')}</span>
          </div>
          <h6 class="card-title fw-bold text-dark mb-2">${escapeHtml(t.title || 'No Title')}</h6>
          <div class="mb-3">
            <small class="text-muted d-block mb-1">Category: <strong class="text-dark">${escapeHtml(t.category || 'Work-related')}</strong></small>
            <small class="text-muted d-block">Due: <strong class="text-dark">${dueFormatted || '—'}</strong> ${overdueBadge}</small>
            ${completedFormatted ? `<small class="text-muted d-block mt-1">Completed: <strong class="text-success">${completedFormatted}</strong></small>` : ''}
          </div>
          <div class="mb-3">
            <div class="d-flex justify-content-between align-items-center mb-1">
              <span class="text-muted small">Progress</span>
              <span class="fw-bold small text-primary">${prog}%</span>
            </div>
            <div class="progress" style="height:8px;">
              <div class="progress-bar" role="progressbar" style="width:${prog}%;" aria-valuenow="${prog}" aria-valuemin="0" aria-valuemax="100"></div>
            </div>
          </div>
          <div class="d-flex justify-content-end border-top pt-2 mt-2">
            ${actions}
          </div>
        </div>
      </div>
    `;
  }

  window.editPersonalTask = function (id) {
    const t = personalTasks.find(x => x.id === id);
    if (!t) return;
    const f = elements.personalTaskForm; if (!f) return;
    f.querySelector('#ptmId').value = t.id;
    f.querySelector('#ptmTitle').value = t.title || '';
    f.querySelector('#ptmCategory').value = t.category || 'Work-related';
    f.querySelector('#ptmPriority').value = t.priority || 'Medium';
    f.querySelector('#ptmDue').value = normalizeDateStr(t.due || '');
    try { const td = f.querySelector('#ptmTaskDate'); if (td) td.value = normalizeDateStr(t.taskDate || ''); } catch (_) { }
    f.querySelector('#ptmStatus').value = t.status || 'Not Started';
    f.querySelector('#ptmDesc').value = t.description || '';
    elements.ptmProgress.value = Number(t.progress || 0);
    elements.ptmProgressLabel.textContent = `${elements.ptmProgress.value}%`;
    if ((t.status || '') === 'Done') { elements.ptmProgress.value = 100; elements.ptmProgressLabel.textContent = '100%'; }
    const dcomp = normalizeDateStr(t.dateCompleted || '');
    const dcEl = f.querySelector('#ptmDateCompleted'); if (dcEl) { dcEl.value = dcomp; dcEl.required = ((t.status || '') === 'Done'); }
    // Subtasks
    renderSubtasks(t.subtasks || []);
    elements.personalTaskModal.show();
  };

  window.deletePersonalTask = async function (id) {
    if (!confirm('Delete this task?')) return;
    try { await personalTaskService.remove(id); }
    catch (err) { alert(err.message); }
  };

  function renderSubtasks(list) {
    const wrap = elements.ptmSubtasks; if (!wrap) return;
    wrap.innerHTML = '';
    (list || []).forEach(item => addSubtaskRow(item));
    if (list.length === 0) addSubtaskRow();
  }

  function addSubtaskRow(data) {
    const wrap = elements.ptmSubtasks; if (!wrap) return;
    const row = document.createElement('div');
    row.className = 'd-flex align-items-center gap-2';
    row.innerHTML = `<input type="checkbox" class="form-check-input ptm-subtask-done" ${data?.done ? 'checked' : ''} />`
      + `<input type="text" class="form-control form-control-sm ptm-subtask-text" placeholder="Subtask" value="${escapeAttr(data?.text || '')}" />`
      + `<button class="btn btn-sm btn-outline-danger"><i class="fa fa-trash"></i></button>`;
    row.querySelector('button').onclick = () => { row.remove(); };
    wrap.appendChild(row);
  }

  function collectSubtasks() {
    const wrap = elements.ptmSubtasks; if (!wrap) return [];
    return Array.from(wrap.querySelectorAll('.d-flex')).map(div => ({
      done: div.querySelector('.ptm-subtask-done')?.checked || false,
      text: div.querySelector('.ptm-subtask-text')?.value?.trim() || ''
    })).filter(x => x.text);
  }
  window.renderPersonalTasks = renderPersonalTasks;

  function getFilteredPersonalTasks() {
    let list = personalTasks.slice();
    const cat = elements.ptmCategoryFilter?.value || '';
    const stat = elements.ptmStatusFilter?.value || '';
    const prio = elements.ptmPriorityFilter?.value || '';
    const text = (elements.ptmTitleSearch?.value || '').toLowerCase().trim();
    if (cat) list = list.filter(t => (t.category || '') === cat);
    if (stat) list = list.filter(t => (t.status || '') === stat);
    if (prio) list = list.filter(t => (t.priority || '') === prio);
    if (text) list = list.filter(t => (t.title || '').toLowerCase().includes(text));
    // sort: unfinished first (due-today, then overdue, then upcoming, then no-due); Done at bottom
    const today = fmtYmd(new Date());
    const prioRank = (p) => ({ 'High': 0, 'Medium': 1, 'Low': 2 }[p] ?? 3);
    list.sort((a, b) => {
      const aDone = (a.status || '') === 'Done';
      const bDone = (b.status || '') === 'Done';
      if (aDone !== bDone) return aDone ? 1 : -1; // push Done to bottom

      const aDue = normalizeYmd(a.due || '');
      const bDue = normalizeYmd(b.due || '');
      const aToday = aDue && safeDateCompare(aDue, today) === 0;
      const bToday = bDue && safeDateCompare(bDue, today) === 0;
      if (aToday !== bToday) return aToday ? -1 : 1;

      const aOver = aDue && safeDateCompare(aDue, today) < 0;
      const bOver = bDue && safeDateCompare(bDue, today) < 0;
      if (aOver !== bOver) return aOver ? -1 : 1;

      const aHasDue = !!aDue;
      const bHasDue = !!bDue;
      if (aHasDue !== bHasDue) return aHasDue ? -1 : 1;

      // tie-breakers among same bucket: priority then due date
      const pr = prioRank(a.priority || '') - prioRank(b.priority || '');
      if (pr !== 0) return pr;

      // for Done group, prefer most recently completed first
      if (aDone && bDone) {
        const ac = normalizeYmd(a.dateCompleted || '');
        const bc = normalizeYmd(b.dateCompleted || '');
        if (ac && bc) return safeDateCompare(bc, ac); // newer first
        if (ac || bc) return ac ? -1 : 1;
      }

      // default: by due date asc if present, else by title
      if (aDue || bDue) return (aDue || '').localeCompare(bDue || '');
      return (a.title || '').localeCompare(b.title || '');
    });
    return list;
  }

  function renderPersonalTasks() {
    const tbody = elements.personalTasksTbody; if (!tbody) return;
    const list = getFilteredPersonalTasks();
    // counters
    const todo = list.filter(t => (t.status || '') === 'Not Started').length;
    const ongoing = list.filter(t => (t.status || '') === 'Ongoing').length;
    const blocked = list.filter(t => (t.status || '') === 'Blocked').length;
    const done = list.filter(t => (t.status || '') === 'Done').length;
    if (elements.ptmCountTodo) elements.ptmCountTodo.textContent = todo;
    if (elements.ptmCountOngoing) elements.ptmCountOngoing.textContent = ongoing;
    if (elements.ptmCountBlocked) elements.ptmCountBlocked.textContent = blocked;
    if (elements.ptmCountDone) elements.ptmCountDone.textContent = done;
    // pagination
    const total = list.length;
    const per = PERSONAL_TASKS_PER_PAGE;
    const maxPage = Math.max(1, Math.ceil(total / per));
    if (personalTasksPage > maxPage) personalTasksPage = maxPage;
    const start = (personalTasksPage - 1) * per;
    const end = Math.min(start + per, total);
    const slice = list.slice(start, end);
    tbody.innerHTML = slice.map(personalTaskRowHtml).join('');
    if (elements.personalTasksMobileList) {
      elements.personalTasksMobileList.innerHTML = slice.map(personalTaskCardHtml).join('');
    }
    attachPersonalTaskRowEvents();
  }

  function attachPersonalTaskRowEvents() {
    try {
      document.querySelectorAll('#personalTasksTable tbody tr').forEach(tr => {
        tr.style.cursor = 'pointer';
        tr.onclick = (e) => {
          // Ignore clicks on action buttons
          if (e.target.closest('button')) return;
          viewPersonalTask(tr.dataset.id);
        };
      });
      document.querySelectorAll('#personalTasksMobileList .card').forEach(card => {
        card.onclick = (e) => {
          // Ignore clicks on action buttons
          if (e.target.closest('button')) return;
          viewPersonalTask(card.dataset.id);
        };
      });
    } catch (_) {/* noop */ }
  }

  function viewPersonalTask(id) {
    const t = personalTasks.find(x => x.id === id); if (!t) return;
    const body = elements.ptmDetailsBody; if (!body) return;
    const fmtUi = window.AppUtils && window.AppUtils.formatDateUI ? window.AppUtils.formatDateUI : (x=>x);
    const due = fmtUi(normalizeDateStr(t.due || ''));
    const taskDate = fmtUi(normalizeDateStr(t.taskDate || ''));
    const dateCompleted = fmtUi(normalizeDateStr(t.dateCompleted || ''));
    // Compute duration as WORK DAYS (Mon-Fri) from taskDate to dateCompleted.
    // Rule: if completed within the same day, count as 1 work day.
    let durationText = '-';
    try {
      function countWeekdaysInclusive(sdStr, edStr) {
        const s = new Date(sdStr + 'T00:00:00');
        const e = new Date(edStr + 'T00:00:00');
        if (!isFinite(s) || !isFinite(e) || e < s) return 0;
        let cnt = 0;
        for (const d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
          const dow = d.getDay(); // 0=Sun, 6=Sat
          if (dow !== 0 && dow !== 6) cnt++;
        }
        return cnt;
      }
      if (t.taskDate && t.dateCompleted) {
        if (t.taskDate === t.dateCompleted) {
          durationText = '1 work day';
        } else {
          const days = countWeekdaysInclusive(t.taskDate, t.dateCompleted);
          durationText = `${days} work day${days === 1 ? '' : 's'}`;
        }
      }
    } catch (_) { durationText = '-'; }
    const prog = (t.status === 'Done') ? 100 : Math.max(0, Math.min(100, Number(t.progress || 0)));
    const subt = (t.subtasks || []);
    const subtRows = subt.length ? subt.map(s => `<div class="d-flex align-items-center gap-2 small">
        <i class="fa ${s.done ? 'fa-check-circle text-success' : 'fa-circle text-muted'}"></i>
        <span>${escapeHtml(s.text || '')}</span>
      </div>`).join('') : '<div class="text-muted small">No subtasks</div>';
    body.innerHTML = `
      <div class="row g-2 mb-2">
        <div class="col-md-8">
          <h5 class="fw-bold mb-1">${escapeHtml(t.title || '')}</h5>
          ${t.description ? `<div class="small">${escapeHtml(t.description)}</div>` : ''}
        </div>
        <div class="col-md-4">
          <div class="small text-muted">Status</div>
          <div class="fw-semibold">${escapeHtml(t.status || '')}</div>
        </div>
      </div>
      <div class="row g-2 mb-2">
        <div class="col-md-3">
          <div class="small text-muted">Category</div>
          <div class="fw-semibold">${escapeHtml(t.category || '')}</div>
        </div>
        <div class="col-md-3">
          <div class="small text-muted">Priority</div>
          <div class="fw-semibold">${escapeHtml(t.priority || '')}</div>
        </div>
        <div class="col-md-3">
          <div class="small text-muted">Due Date</div>
          <div class="fw-semibold">${due || '-'}</div>
        </div>
        <div class="col-md-3">
          <div class="small text-muted">Date Completed</div>
          <div class="fw-semibold">${dateCompleted || '-'}</div>
        </div>
      </div>
      <div class="row g-2 mb-2">
        <div class="col-md-3">
          <div class="small text-muted">Task Date</div>
          <div class="fw-semibold">${taskDate || '-'}</div>
        </div>
        <div class="col-md-3">
          <div class="small text-muted">Duration</div>
          <div class="fw-semibold">${durationText}</div>
        </div>
      </div>
      <div class="mb-2">
        <div class="small text-muted mb-1">Progress</div>
        <div class="progress" style="height:10px;">
          <div class="progress-bar" role="progressbar" style="width:${prog}%" aria-valuenow="${prog}" aria-valuemin="0" aria-valuemax="100"></div>
        </div>
      </div>
      <div class="mt-3">
        <div class="small text-muted mb-1">Subtasks</div>
        <div>${subtRows}</div>
      </div>
    `;
    elements.personalTaskDetailsModal.show();
  }

  function clearPersonalTaskForm() {
    const f = elements.personalTaskForm; if (!f) return;
    f.reset();
    f.querySelector('#ptmId').value = '';
    if (elements.ptmProgress) { elements.ptmProgress.value = 0; elements.ptmProgressLabel.textContent = '0%'; }
    try { const dc = f.querySelector('#ptmDateCompleted'); if (dc) { dc.value = ''; dc.required = false; } } catch (_) { }
    try { const td = f.querySelector('#ptmTaskDate'); if (td) { td.value = new Date().toISOString().slice(0, 10); } } catch (_) { }
    renderSubtasks([]);
  }

  async function savePersonalTask(e) {
    e.preventDefault();
    const user = auth.currentUser;
    if (!user) { alert('Please login'); return; }
    const f = elements.personalTaskForm; if (!f) return;
    const id = f.querySelector('#ptmId').value || (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36));
    const payload = {
      id,
      ownerUid: user.uid,
      ownerEmail: user.email || '',
      title: f.querySelector('#ptmTitle').value.trim(),
      category: f.querySelector('#ptmCategory').value,
      priority: f.querySelector('#ptmPriority').value,
      due: normalizeDateStr(f.querySelector('#ptmDue').value),
      taskDate: normalizeDateStr(f.querySelector('#ptmTaskDate')?.value || ''),
      status: f.querySelector('#ptmStatus').value,
      description: f.querySelector('#ptmDesc').value.trim(),
      progress: Number(elements.ptmProgress?.value || 0),
      dateCompleted: normalizeDateStr(f.querySelector('#ptmDateCompleted')?.value || ''),
      subtasks: collectSubtasks(),
      updatedAt: serverTimestamp(),
      createdAt: serverTimestamp(),
    };
    if ((payload.status || '') === 'Done') {
      payload.progress = 100;
      if (!payload.dateCompleted) { payload.dateCompleted = new Date().toISOString().slice(0, 10); }
    } else {
      // If not Done, clear Date Completed
      payload.dateCompleted = '';
    }
    if (!payload.title) { alert('Title is required'); return; }
    if (payload.ownerUid !== user.uid) {
      console.error('[PTM] ownerUid mismatch', { payloadOwner: payload.ownerUid, authUid: user.uid });
      alert('Cannot save: owner mismatch');
      return;
    }
    console.log('[PTM] saving task', payload);
    try {
      await personalTaskService.collection().doc(id).set(payload, { merge: true });
      elements.personalTaskModal.hide();
      clearPersonalTaskForm();
    } catch (err) {
      console.error('[PTM] save error', err);
      if (err && err.code === 'permission-denied') {
        alert('Missing or insufficient permissions. Please ensure Firestore rules allow creating personal tasks for your account.');
      } else {
        alert(err.message || String(err));
      }
    }
  }

  function subscribePersonalTasks() {
    const user = auth.currentUser;
    if (unsubscribePersonalTasks) { unsubscribePersonalTasks(); unsubscribePersonalTasks = null; }
    if (!user) { personalTasks = []; renderPersonalTasks(); return; }
    unsubscribePersonalTasks = personalTaskService.collection()
      .where('ownerUid', '==', user.uid)
      .onSnapshot(snap => {
        personalTasks = snap.docs.map(d => ({ id: d.id, ...d.data() }));
        PersonalTasksFeatureInstance.setPersonalTasks(personalTasks);
        renderPersonalTasks();
      });
  }
  let unsubscribeProjects = null;
  let unsubscribeDeepwells = null;
  let unsubscribeServiceUpdates = null;
  let unsubscribePresentations = null;

  let isAdmin = false;

  const elements = {
    projectsContainer: document.getElementById("projectsContainer"),
    projectsPagination: document.getElementById('projectsPagination'),
    agencyFilter: document.getElementById("agencyFilter"),
    statusFilter: document.getElementById("statusFilter"),
    searchInput: document.getElementById("searchInput"),
    projectForm: document.getElementById("projectForm"),
    projectModal: (function () { const el = document.getElementById('projectModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    exportBtn: document.getElementById("exportBtn"),
    loginForm: document.getElementById('loginForm'),
    signupForm: document.getElementById('signupForm'),
    showSignupLink: document.getElementById('showSignupLink'),
    showLoginLink: document.getElementById('showLoginLink'),
    loginLinks: document.getElementById('loginLinks'),
    signupLinks: document.getElementById('signupLinks'),
    manageUsersBtn: document.getElementById('manageUsersBtn'),
    pendingUsersTable: document.getElementById('pendingUsersTable'),
    approvedUsersTable: document.getElementById('approvedUsersTable'),
    // Deepwell specific
    deepwellsTbody: document.getElementById('deepwellsTbody'),
    deepwellsPagination: document.getElementById('deepwellsPagination'),
    dwProviderFilter: document.getElementById('dwProviderFilter'),
    dwStatusFilter: document.getElementById('dwStatusFilter'),
    dwProdSort: document.getElementById('dwProdSort'),
    dwSearchInput: document.getElementById('dwSearchInput'),
    deepwellForm: document.getElementById('deepwellForm'),
    deepwellModal: (function () { const el = document.getElementById('deepwellModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    generateDeepwellReportBtn: document.getElementById('generateDeepwellReportBtn'),
    projectsTab: document.getElementById('projectsTab'),
    deepwellsTab: document.getElementById('deepwellsTab'),
    projectsSection: document.getElementById('projectsSection'),
    deepwellsSection: document.getElementById('deepwellsSection'),
    // Reforestation specific
    reforestationTab: document.getElementById('reforestationTab'),
    reforestationSection: document.getElementById('reforestationSection'),
    reforestationTbody: document.getElementById('reforestationTbody'),
    reforestationTypeFilter: document.getElementById('reforestationTypeFilter'),
    reforestationStatusFilter: document.getElementById('reforestationStatusFilter'),
    reforestationSearchInput: document.getElementById('reforestationSearchInput'),
    addReforestationBtn: document.getElementById('addReforestationBtn'),
    reforestationModal: (function () { const el = document.getElementById('reforestationModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    reforestationForm: document.getElementById('reforestationForm'),
    addDeepwellBtn: document.getElementById('addDeepwellBtn'),
    deepwellsTbody: document.getElementById('deepwellsTbody'),
    dwMonthsBody: document.getElementById('dwMonthsBody'),
    addDwMonthBtn: document.getElementById('addDwMonthBtn'),
    // Deepwell monthly chart
    deepwellMonthlyChartCanvas: document.getElementById('deepwellMonthlyChart'),
    dwChartEmpty: document.getElementById('dwChartEmpty'),
    // Reforestation details modal
    reforestationDetailsBody: document.getElementById('reforestationDetailsBody'),
    reforestationDetailsModal: (function () { const el = document.getElementById('reforestationDetailsModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    // Service Update Report
    serviceUpdateTab: document.getElementById('serviceUpdateTab'),
    serviceUpdateSection: document.getElementById('serviceUpdateSection'),
    serviceStartDateFilter: document.getElementById('serviceStartDateFilter'),
    serviceEndDateFilter: document.getElementById('serviceEndDateFilter'),
    serviceProviderFilter: document.getElementById('serviceProviderFilter'),
    serviceYearFilter: document.getElementById('serviceYearFilter'),
    serviceUpdatesTbody: document.getElementById('serviceUpdatesTbody'),
    addServiceUpdateBtn: document.getElementById('addServiceUpdateBtn'),
    surBulkTemplateBtn: document.getElementById('surBulkTemplateBtn'),
    surBulkImportBtn: document.getElementById('surBulkImportBtn'),
    surBulkImportFile: document.getElementById('surBulkImportFile'),
    surExportBtn: document.getElementById('surExportBtn'),
    serviceUpdateModal: (function () { const el = document.getElementById('serviceUpdateModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    serviceUpdateForm: document.getElementById('serviceUpdateForm'),
    serviceUpdateDetailsModal: (function () { const el = document.getElementById('serviceUpdateDetailsModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    // Service Update chart elements
    surChartCanvas: document.getElementById('surChart'),
    surChartEmpty: document.getElementById('surChartEmpty'),
    surChartView: document.getElementById('surChartView'),
    // Service Update pagination controls
    serviceUpdatesPageInfo: document.getElementById('serviceUpdatesPageInfo'),
    serviceUpdatesPrev: document.getElementById('serviceUpdatesPrev'),
    serviceUpdatesNext: document.getElementById('serviceUpdatesNext'),
    // Product Presentations
    presentationsTab: document.getElementById('presentationsTab'),
    presentationsSection: document.getElementById('presentationsSection'),
    calendarTab: document.getElementById('calendarTab'),
    calendarSection: document.getElementById('calendarSection'),
    presentationStartDateFilter: document.getElementById('presentationStartDateFilter'),
    presentationEndDateFilter: document.getElementById('presentationEndDateFilter'),
    presentationYearFilter: document.getElementById('presentationYearFilter'),
    presentationSubjectFilter: document.getElementById('presentationSubjectFilter'),
    presentationPresenterFilter: document.getElementById('presentationPresenterFilter'),
    presentationsTbody: document.getElementById('presentationsTbody'),
    presentationsMobileList: document.getElementById('presentationsMobileList'),
    presentationsEmpty: document.getElementById('presentationsEmpty'),
    // Personal Tasks
    documentRegistryBtn: document.getElementById('documentRegistryBtn'),
    documentRegistrySection: document.getElementById('documentRegistrySection'),
    personalTasksBtn: document.getElementById('personalTasksBtn'),
    personalTasksSection: document.getElementById('personalTasksSection'),
    personalTasksTbody: document.getElementById('personalTasksTbody'),
    personalTasksMobileList: document.getElementById('personalTasksMobileList'),
    ptmTitleSearch: document.getElementById('ptmTitleSearch'),
    ptmAddBtn: document.getElementById('ptmAddBtn'),
    ptmCategoryFilter: document.getElementById('ptmCategoryFilter'),
    ptmStatusFilter: document.getElementById('ptmStatusFilter'),
    ptmPriorityFilter: document.getElementById('ptmPriorityFilter'),
    ptmCountTodo: document.getElementById('ptmCountTodo'),
    ptmCountOngoing: document.getElementById('ptmCountOngoing'),
    ptmCountBlocked: document.getElementById('ptmCountBlocked'),
    ptmCountDone: document.getElementById('ptmCountDone'),
    personalTaskForm: document.getElementById('personalTaskForm'),
    personalTaskModal: (function () { const el = document.getElementById('personalTaskModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    ptmProgress: document.getElementById('ptmProgress'),
    ptmProgressLabel: document.getElementById('ptmProgressLabel'),
    ptmSubtasks: document.getElementById('ptmSubtasks'),
    ptmAddSubtask: document.getElementById('ptmAddSubtask'),
    // Personal Task Details (read-only modal)
    ptmDetailsBody: document.getElementById('ptmDetailsBody'),
    personalTaskDetailsModal: (function () { const el = document.getElementById('personalTaskDetailsModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    addPresentationBtn: document.getElementById('addPresentationBtn'),
    presentationModal: (function () { const el = document.getElementById('presentationModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    presentationForm: document.getElementById('presentationForm'),
    presentationId: document.getElementById('presentationId'),
    presentationSubject: document.getElementById('presentationSubject'),
    presentationVenue: document.getElementById('presentationVenue'),
    presentationDate: document.getElementById('presentationDate'),
    presentationTime: document.getElementById('presentationTime'),
    presentationPresenter: document.getElementById('presentationPresenter'),
    presentationRemarks: document.getElementById('presentationRemarks'),
    presentationYearLabel: document.getElementById('presentationYearLabel'),

    // OPCR Dashboard
    headerOpcrBtn: document.getElementById('headerOpcrBtn'),
    opcrSection: document.getElementById('opcrSection'),
    opcrTbody: document.getElementById('opcrTbody'),
    opcrEmpty: document.getElementById('opcrEmpty'),
    opcrYearFilter: document.getElementById('opcrYearFilter'),
    opcrSearchInput: document.getElementById('opcrSearchInput'),
    addOpcrBtn: document.getElementById('addOpcrBtn'),
    opcrModal: (function () { const el = document.getElementById('opcrModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    opcrForm: document.getElementById('opcrForm'),
    opcrId: document.getElementById('opcrId'),
    opcrDepartment: document.getElementById('opcrDepartment'),
    opcrYear: document.getElementById('opcrYear'),
    opcrPeriodStart: document.getElementById('opcrPeriodStart'),
    opcrPeriodEnd: document.getElementById('opcrPeriodEnd'),
    opcrStatus: document.getElementById('opcrStatus'),
    opcrRemarks: document.getElementById('opcrRemarks'),
    opcrFormSection: document.getElementById('opcrFormSection'),

    // IPCR Dashboard
    headerIpcrBtn: document.getElementById('headerIpcrBtn'),
    ipcrSection: document.getElementById('ipcrSection'),
    ipcrTbody: document.getElementById('ipcrTbody'),
    ipcrEmpty: document.getElementById('ipcrEmpty'),
    ipcrYearFilter: document.getElementById('ipcrYearFilter'),
    ipcrSearchInput: document.getElementById('ipcrSearchInput'),
    addIpcrBtn: document.getElementById('addIpcrBtn'),
    ipcrModal: (function () { const el = document.getElementById('ipcrModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    ipcrForm: document.getElementById('ipcrForm'),
    ipcrId: document.getElementById('ipcrId'),
    ipcrFormRateeName: document.getElementById('ipcrFormRateeName'),
    ipcrFormRateePosition: document.getElementById('ipcrFormRateePosition'),
    ipcrFormDepartment: document.getElementById('ipcrFormDepartment'),
    ipcrFormYear: document.getElementById('ipcrFormYear'),
    ipcrFormPeriodStart: document.getElementById('ipcrFormPeriodStart'),
    ipcrFormPeriodEnd: document.getElementById('ipcrFormPeriodEnd'),
    generateIpcrBtn: document.getElementById('generateIpcrBtn'),
    generateIpcrModal: (function () { const el = document.getElementById('generateIpcrModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),

    // Dashboard (desktop/mobile)
    dashboardTab: document.getElementById('dashboardTab'),
    dashboardSection: document.getElementById('dashboardSection'),
    dashProjectsTotal: document.getElementById('dashProjectsTotal'),
    dashProjectsCompletedRecently: document.getElementById('dashProjectsCompletedRecently'),
    dashDeepwellsTotal: document.getElementById('dashDeepwellsTotal'),
    dashReforestationsTotal: document.getElementById('dashReforestationsTotal'),
    // Dashboard breakdowns
    dashProjectsOngoing: document.getElementById('dashProjectsOngoing'),
    dashProjectsCompleted: document.getElementById('dashProjectsCompleted'),
    dashProjectsMWCI: document.getElementById('dashProjectsMWCI'),
    dashProjectsMWSI: document.getElementById('dashProjectsMWSI'),
    dashProjectsMWSS: document.getElementById('dashProjectsMWSS'),
    dashDeepwellsActive: document.getElementById('dashDeepwellsActive'),
    dashDeepwellsInactive: document.getElementById('dashDeepwellsInactive'),
    dashDeepwellsStandby: document.getElementById('dashDeepwellsStandby'),
    // Deepwell YTD Production
    dashDwYtdYear: document.getElementById('dashDwYtdYear'),
    dashDwYtdMwci: document.getElementById('dashDwYtdMwci'),
    dashDwYtdMwsi: document.getElementById('dashDwYtdMwsi'),
    // Dam elevations (today/latest)
    dashAngat: document.getElementById('dashAngat'),
    dashIpo: document.getElementById('dashIpo'),
    dashLaMesa: document.getElementById('dashLaMesa'),
    dashDamAsOf: document.getElementById('dashDamAsOf'),
    // Dam elevations YTD highs/lows (new detailed layout)
    dashAngatYtdHiVal: document.getElementById('dashAngatYtdHiVal'),
    dashAngatYtdHiDate: document.getElementById('dashAngatYtdHiDate'),
    dashAngatYtdLoVal: document.getElementById('dashAngatYtdLoVal'),
    dashAngatYtdLoDate: document.getElementById('dashAngatYtdLoDate'),
    dashIpoYtdHiVal: document.getElementById('dashIpoYtdHiVal'),
    dashIpoYtdHiDate: document.getElementById('dashIpoYtdHiDate'),
    dashIpoYtdLoVal: document.getElementById('dashIpoYtdLoVal'),
    dashIpoYtdLoDate: document.getElementById('dashIpoYtdLoDate'),
    dashLaMesaYtdHiVal: document.getElementById('dashLaMesaYtdHiVal'),
    dashLaMesaYtdHiDate: document.getElementById('dashLaMesaYtdHiDate'),
    dashLaMesaYtdLoVal: document.getElementById('dashLaMesaYtdLoVal'),
    dashLaMesaYtdLoDate: document.getElementById('dashLaMesaYtdLoDate'),
    // Backward-compat (old single-line YTD fields, if present)
    dashAngatYtd: document.getElementById('dashAngatYtd'),
    dashIpoYtd: document.getElementById('dashIpoYtd'),
    dashLaMesaYtd: document.getElementById('dashLaMesaYtd'),
    // Deprecated (card removed): keep for backward safety if present
    dashSurNetSupply: document.getElementById('dashSurNetSupply'),
    surChartDashCanvas: document.getElementById('surChartDash'),
    dashSurChartEmpty: document.getElementById('dashSurChartEmpty'),
    dwMonthlyChartDashCanvas: document.getElementById('dwMonthlyChartDash'),
    dashDwChartEmpty: document.getElementById('dashDwChartEmpty'),
    deepwellMapCard: document.getElementById('deepwellMapCard'),
    deepwellMapPreview: document.getElementById('deepwellMapPreview'),
    dashDeepwellMapEmpty: document.getElementById('dashDeepwellMapEmpty'),
    deepwellMapDetail: document.getElementById('deepwellMapDetail'),
    deepwellMapModalEmpty: document.getElementById('deepwellMapModalEmpty'),

    // Dashboard Presentations chart
    presentationChartDashCanvas: document.getElementById('presentationChartDash'),
    dashPresentationChartEmpty: document.getElementById('dashPresentationChartEmpty'),
    damElevationsChartDashCanvas: document.getElementById('damElevationsChartDash'),
    dashDamChartEmpty: document.getElementById('dashDamChartEmpty'),
    treatmentPlantsMwciChartDashCanvas: document.getElementById('treatmentPlantsMwciChartDash'),
    dashTreatmentPlantsMwciEmpty: document.getElementById('dashTreatmentPlantsMwciEmpty'),
    treatmentPlantsMwsiChartDashCanvas: document.getElementById('treatmentPlantsMwsiChartDash'),
    dashTreatmentPlantsMwsiEmpty: document.getElementById('dashTreatmentPlantsMwsiEmpty'),
    treatmentPlantsExpandedModal: (function () { const el = document.getElementById('treatmentPlantsExpandedModal'); return el ? new bootstrap.Modal(el) : { show() { }, hide() { } }; })(),
    treatmentPlantsExpandedTitle: document.getElementById('treatmentPlantsExpandedTitle'),
    treatmentPlantsExpandedSubtitle: document.getElementById('treatmentPlantsExpandedSubtitle'),
    treatmentPlantsExpandedSummary: document.getElementById('treatmentPlantsExpandedSummary'),
    treatmentPlantsExpandedChartCanvas: document.getElementById('treatmentPlantsExpandedChart'),
    treatmentPlantsExpandedTbody: document.getElementById('treatmentPlantsExpandedTbody'),
    treatmentPlantsExpandedEmpty: document.getElementById('treatmentPlantsExpandedEmpty'),
    treatmentPlantsPrintBtn: document.getElementById('treatmentPlantsPrintBtn'),
    treatmentPlantsWordBtn: document.getElementById('treatmentPlantsWordBtn'),
  };

  // convenient references
  const { loginForm, signupForm, showSignupLink, showLoginLink, loginLinks, signupLinks } = elements;
  const appWrapper = document.getElementById('appWrapper');
  const sidebarToggle = document.getElementById('sidebarToggle');
  const billingContainer = document.getElementById('billingContainer');
  const addBillingBtn = document.getElementById('addBillingBtn');
  const viewOnlyLink = document.getElementById('viewOnlyLink');
  const logoutBtn = document.getElementById('logoutBtn');
  const loginBtn = document.getElementById('loginBtn');
  const loginScreen = document.getElementById('loginScreen');
  const notificationService = window.NotificationService;
  const notificationsFeature = window.NotificationsFeature;
  if (notificationsFeature) {
    notificationsFeature.init({
      notificationService,
      calendarService,
      auth,
    });
  }

  // Sidebar collapse/expand behavior (init after elements are available)
  try { window.UIShell?.initSidebarToggle(); } catch (_) { /* noop */ }

  // Core UI/access flags
  let isViewOnly = false;
  let isLevel2 = false; // accessLevel 2 users
  let elevatedAccess = false; // isAdmin || isLevel2
  let authStateSeq = 0;

  const projectsFeature = window.ProjectsFeature;
  if (!projectsFeature) {
    throw new Error('ProjectsFeature missing; ensure js/features/projects.js loads before script.js');
  }
  window.__VIEW_ONLY__ = () => isViewOnly;
  const ProjectsFeatureInstance = projectsFeature.init({
    elements,
    projectService,
    auth,
    utils,
    compressImage: typeof compressImage === 'function' ? compressImage : null,
    elevatedAccessRef: () => elevatedAccess,
    isAdminRef: () => isAdmin,
  });
  const deepwellsFeature = window.DeepwellsFeature;
  if (!deepwellsFeature) {
    throw new Error('DeepwellsFeature missing; ensure js/features/deepwells.js loads before script.js');
  }
  const DeepwellsFeatureInstance = deepwellsFeature.init({
    elements,
    deepwellService,
    auth,
    utils,
    perPage: DEEPWELLS_PER_PAGE,
    elevatedAccessRef: () => elevatedAccess,
    isAdminRef: () => isAdmin,
  });
  const serviceUpdatesFeature = window.ServiceUpdatesFeature;
  if (!serviceUpdatesFeature) {
    throw new Error('ServiceUpdatesFeature missing; ensure js/features/service-updates.js loads before script.js');
  }
  const ServiceUpdatesFeatureInstance = serviceUpdatesFeature.init({
    elements,
    serviceUpdateService,
    auth,
    utils,
    perPage: SERVICE_UPDATES_PER_PAGE,
    elevatedAccessRef: () => elevatedAccess,
    isAdminRef: () => isAdmin,
  });
  const presentationsFeature = window.PresentationsFeature;
  if (!presentationsFeature) {
    throw new Error('PresentationsFeature missing; ensure js/features/presentations.js loads before script.js');
  }
  const PresentationsFeatureInstance = presentationsFeature.init({
    elements,
    presentationService,
    calendarService,
    auth,
    utils,
    elevatedAccessRef: () => elevatedAccess,
    isAdminRef: () => isAdmin,
  });
  const personalTasksFeature = window.PersonalTasksFeature;
  if (!personalTasksFeature) {
    throw new Error('PersonalTasksFeature missing; ensure js/features/personal-tasks.js loads before script.js');
  }
  const PersonalTasksFeatureInstance = personalTasksFeature.init({
    elements,
    auth,
    personalTaskService,
    utils,
  });
  const messengerFeature = window.MessengerFeature;
  if (!messengerFeature) {
    throw new Error('MessengerFeature missing; ensure js/features/messenger.js loads before script.js');
  }
  const MessengerFeatureInstance = messengerFeature.init();
  const opcrFeature = window.OpcrFeature;
  if (!opcrFeature) {
    throw new Error('OpcrFeature missing; ensure js/features/opcr.js loads before script.js');
  }
  const OpcrFeatureInstance = opcrFeature.init({
    elements,
    opcrService,
    auth,
    utils,
    isAdminRef: () => isAdmin,
  });

  const calendarFeature = window.CalendarFeature;
  if (!calendarFeature) {
    throw new Error('CalendarFeature missing; ensure js/features/calendar.js loads before script.js');
  }

  const ipcrFeature = window.IpcrFeature;
  if (!ipcrFeature) {
    throw new Error('IpcrFeature missing; ensure js/features/ipcr.js loads before script.js');
  }
  const IpcrFeatureInstance = ipcrFeature.init({
    elements,
    ipcrService,
    opcrService,
    auth,
    utils,
    isAdminRef: () => isAdmin,
  });
  const CalendarFeatureInstance = calendarFeature.init({
    elements,
    calendarService,
    auth,
    utils,
    isAdminRef: () => isAdmin,
  });
  let deepwells = [];
  let reforestations = [];
  let serviceUpdates = [];
  let presentations = [];
  let pendingUsers = [];
  let surChart = null; // Chart.js instance for Service Update (Inflows vs Supply)
  let dashDwMonthlyChart = null; // Dashboard deepwells chart instance
  let dashSurChart = null; // Dashboard SUR chart instance
  let dashPresentationsChart = null; // Dashboard Presentations chart instance
  let dashPresentationsChartSignature = '';

  let dashDeepwellMapPreview = null;
  let dashDeepwellMapDetail = null;
  let dashDeepwellMapPreviewLayer = null;
  let dashDeepwellMapDetailLayer = null;
  let dashDeepwellMapPreviewSignature = '';
  let dashDeepwellMapDetailSignature = '';

  let dashDamChart = null;
  let dashDamChartSignature = '';
  let dashDamYearSelectSignature = '';
  const dashTreatmentPlantCharts = { MWCI: null, MWSI: null };
  const dashTreatmentPlantChartSignatures = { MWCI: '', MWSI: '' };
  const dashTreatmentPlantYearSelectSignatures = { MWCI: '', MWSI: '' };
  let treatmentPlantsExpandedChart = null;
  let treatmentPlantsExpandedSeries = null;


  let approvedUsers = [];
  let personalTasks = [];
  let unsubscribePersonalTasks = null;
  let personalTasksPage = 1;
  let opcrEntries = [];
  let unsubscribeOpcr = null;


  // Construction Projects pagination
  // Deepwells pagination
  // Service Updates pagination
  let serviceUpdatesPage = 1;
  function updatePendingBadge() {
    const badge = document.getElementById('pendingBadge');
    if (!badge) return;
    const cnt = pendingUsers.length;
    if (cnt > 0) { badge.textContent = cnt; badge.style.display = 'inline-block'; }
    else { badge.style.display = 'none'; }
  }

  // Build monthly totals (sum per month) for a specific year (Jan..Dec)
  function buildSurMonthlyTotalsForYear(daily, yearStr) {
    if (!daily || !Array.isArray(daily.labels) || daily.labels.length === 0) {
      const yr = parseInt(String(yearStr || '').slice(0, 4), 10) || new Date().getFullYear();
      const monthsKeys = Array.from({ length: 12 }, (_, i) => `${yr}-${String(i + 1).padStart(2, '0')}`);
      return {
        labels: monthsKeys.map(monthLabel), inflows: Array(12).fill(0), production: Array(12).fill(0),
        productionMWCI: Array(12).fill(0), productionMWSI: Array(12).fill(0), augment: Array(12).fill(0), hasData: false
      };
    }
    let yr = parseInt(String(yearStr || '').slice(0, 4), 10);
    if (isNaN(yr)) {
      const last = daily.labels[daily.labels.length - 1] || '';
      yr = parseInt(last.slice(0, 4), 10);
      if (isNaN(yr)) yr = new Date().getFullYear();
    }
    const monthsKeys = Array.from({ length: 12 }, (_, i) => `${yr}-${String(i + 1).padStart(2, '0')}`);
    const sums = monthsKeys.map(() => ({ inflows: 0, production: 0, productionMWCI: 0, productionMWSI: 0, augment: 0, augmentMWCI: 0, augmentMWSI: 0 }));
    for (let i = 0; i < daily.labels.length; i++) {
      const d = daily.labels[i] || '';
      if (d.slice(0, 4) === String(yr)) {
        const mIdx = parseInt(d.slice(5, 7), 10) - 1;
        if (mIdx >= 0 && mIdx < 12) {
          sums[mIdx].inflows += Number(daily.inflows[i] || 0);
          sums[mIdx].production += Number(daily.production[i] || 0);
          sums[mIdx].productionMWCI += Number(daily.productionMWCI?.[i] || 0);
          sums[mIdx].productionMWSI += Number(daily.productionMWSI?.[i] || 0);
          sums[mIdx].augment += Number(daily.augment[i] || 0);
          sums[mIdx].augmentMWCI += Number(daily.augmentMWCI?.[i] || 0);
          sums[mIdx].augmentMWSI += Number(daily.augmentMWSI?.[i] || 0);
        }
      }
    }
    const hasYearData = daily.labels.some(d => (d || '').slice(0, 4) === String(yr));
    return {
      labels: monthsKeys.map(monthLabel),
      inflows: sums.map(s => +s.inflows.toFixed(2)),
      production: sums.map(s => +s.production.toFixed(2)),
      productionMWCI: sums.map(s => +s.productionMWCI.toFixed(2)),
      productionMWSI: sums.map(s => +s.productionMWSI.toFixed(2)),
      augment: sums.map(s => +s.augment.toFixed(2)),
      augmentMWCI: sums.map(s => +s.augmentMWCI.toFixed(2)),
      augmentMWSI: sums.map(s => +s.augmentMWSI.toFixed(2)),
      hasData: hasYearData
    };
  }

  // ---- Service Update Chart: Raw Inflows (line) vs Net Supply (stacked bars: Production + Augment) ----
  function buildSurDailyAggregation(list) {
    // Aggregate by report date (same day). If missing, fallback to plantsDate.
    const byDate = {};
    list.forEach(s => {
      const dateKey = (s.date || s.plantsDate || '').trim();
      if (!dateKey) return;
      const inflow = Number(s.inflows || 0) || 0;
      const prod = Number(s.production || 0) || 0;
      const aug = Number(s.supplyAug || 0) || 0;
      const prov = (s.provider || '').toUpperCase();
      if (!byDate[dateKey]) byDate[dateKey] = { inflows: 0, production: 0, productionMWCI: 0, productionMWSI: 0, augment: 0, augmentMWCI: 0, augmentMWSI: 0 };
      byDate[dateKey].inflows += inflow;
      byDate[dateKey].production += prod;
      if (prov === 'MWCI') byDate[dateKey].productionMWCI += prod;
      else if (prov === 'MWSI') byDate[dateKey].productionMWSI += prod;

      byDate[dateKey].augment += aug;
      if (prov === 'MWCI') byDate[dateKey].augmentMWCI += aug;
      else if (prov === 'MWSI') byDate[dateKey].augmentMWSI += aug;
    });
    const dates = Object.keys(byDate).sort();
    return {
      labels: dates,
      inflows: dates.map(d => byDate[d].inflows),
      production: dates.map(d => byDate[d].production),
      productionMWCI: dates.map(d => byDate[d].productionMWCI),
      productionMWSI: dates.map(d => byDate[d].productionMWSI),
      augment: dates.map(d => byDate[d].augment),
      augmentMWCI: dates.map(d => byDate[d].augmentMWCI),
      augmentMWSI: dates.map(d => byDate[d].augmentMWSI),
      hasData: dates.length > 0
    };
  }

  function monthLabel(ymd) {
    if (!ymd) return '';
    const [y, m] = ymd.split('-');
    if (!y || !m) return ymd;
    const d = new Date(Number(y), Number(m) - 1, 1);
    return isNaN(d.getTime()) ? `${y}-${m}` : d.toLocaleString('en-US', { month: 'short', year: 'numeric' });
  }

  function buildSurMonthlyFromDaily(daily) {
    // daily: {labels: [YYYY-MM-DD], inflows[], production[], productionMWCI[], productionMWSI[], augment[], augmentMWCI[], augmentMWSI[]}
    const byMonth = {};
    daily.labels.forEach((d, idx) => {
      const key = (d || '').slice(0, 7); // YYYY-MM
      if (!key) return;
      if (!byMonth[key]) byMonth[key] = { inflows: 0, production: 0, productionMWCI: 0, productionMWSI: 0, augment: 0, augmentMWCI: 0, augmentMWSI: 0, days: 0 };
      byMonth[key].inflows += Number(daily.inflows[idx] || 0);
      byMonth[key].production += Number(daily.production[idx] || 0);
      byMonth[key].productionMWCI += Number(daily.productionMWCI?.[idx] || 0);
      byMonth[key].productionMWSI += Number(daily.productionMWSI?.[idx] || 0);
      byMonth[key].augment += Number(daily.augment[idx] || 0);
      byMonth[key].augmentMWCI += Number(daily.augmentMWCI?.[idx] || 0);
      byMonth[key].augmentMWSI += Number(daily.augmentMWSI?.[idx] || 0);
      byMonth[key].days += 1; // count unique days already consolidated
    });
    const months = Object.keys(byMonth).sort();
    return {
      labels: months.map(monthLabel),
      inflows: months.map(k => byMonth[k].days ? +(byMonth[k].inflows / byMonth[k].days).toFixed(2) : 0),
      production: months.map(k => byMonth[k].days ? +(byMonth[k].production / byMonth[k].days).toFixed(2) : 0),
      productionMWCI: months.map(k => byMonth[k].days ? +(byMonth[k].productionMWCI / byMonth[k].days).toFixed(2) : 0),
      productionMWSI: months.map(k => byMonth[k].days ? +(byMonth[k].productionMWSI / byMonth[k].days).toFixed(2) : 0),
      augment: months.map(k => byMonth[k].days ? +(byMonth[k].augment / byMonth[k].days).toFixed(2) : 0),
      augmentMWCI: months.map(k => byMonth[k].days ? +(byMonth[k].augmentMWCI / byMonth[k].days).toFixed(2) : 0),
      augmentMWSI: months.map(k => byMonth[k].days ? +(byMonth[k].augmentMWSI / byMonth[k].days).toFixed(2) : 0),
      hasData: months.length > 0
    };
  }

  // Build yearly totals across available data (sum per year)
  function buildSurYearlyFromDaily(daily, mode = 'avg') {
    // daily: {labels: [YYYY-MM-DD], inflows[], production[], productionMWCI[], productionMWSI[], augment[], augmentMWCI[], augmentMWSI[]}
    if (!daily || !Array.isArray(daily.labels)) {
      return { labels: [], inflows: [], production: [], productionMWCI: [], productionMWSI: [], augment: [], augmentMWCI: [], augmentMWSI: [], hasData: false };
    }
    const byYear = {};
    daily.labels.forEach((d, idx) => {
      const y = (d || '').slice(0, 4);
      if (!y) return;
      if (!byYear[y]) byYear[y] = { inflows: 0, production: 0, productionMWCI: 0, productionMWSI: 0, augment: 0, augmentMWCI: 0, augmentMWSI: 0, days: 0 };
      byYear[y].inflows += Number(daily.inflows[idx] || 0);
      byYear[y].production += Number(daily.production[idx] || 0);
      byYear[y].productionMWCI += Number(daily.productionMWCI?.[idx] || 0);
      byYear[y].productionMWSI += Number(daily.productionMWSI?.[idx] || 0);
      byYear[y].augment += Number(daily.augment[idx] || 0);
      byYear[y].augmentMWCI += Number(daily.augmentMWCI?.[idx] || 0);
      byYear[y].augmentMWSI += Number(daily.augmentMWSI?.[idx] || 0);
      byYear[y].days += 1;
    });
    const years = Object.keys(byYear).sort();
    const isSum = (String(mode || '').toLowerCase() === 'sum');
    return {
      labels: years,
      inflows: years.map(y => isSum ? +byYear[y].inflows.toFixed(2) : (byYear[y].days ? +(byYear[y].inflows / byYear[y].days).toFixed(2) : 0)),
      production: years.map(y => isSum ? +byYear[y].production.toFixed(2) : (byYear[y].days ? +(byYear[y].production / byYear[y].days).toFixed(2) : 0)),
      productionMWCI: years.map(y => isSum ? +byYear[y].productionMWCI.toFixed(2) : (byYear[y].days ? +(byYear[y].productionMWCI / byYear[y].days).toFixed(2) : 0)),
      productionMWSI: years.map(y => isSum ? +byYear[y].productionMWSI.toFixed(2) : (byYear[y].days ? +(byYear[y].productionMWSI / byYear[y].days).toFixed(2) : 0)),
      augment: years.map(y => isSum ? +byYear[y].augment.toFixed(2) : (byYear[y].days ? +(byYear[y].augment / byYear[y].days).toFixed(2) : 0)),
      augmentMWCI: years.map(y => isSum ? +byYear[y].augmentMWCI.toFixed(2) : (byYear[y].days ? +(byYear[y].augmentMWCI / byYear[y].days).toFixed(2) : 0)),
      augmentMWSI: years.map(y => isSum ? +byYear[y].augmentMWSI.toFixed(2) : (byYear[y].days ? +(byYear[y].augmentMWSI / byYear[y].days).toFixed(2) : 0)),
      hasData: years.length > 0
    };
  }

  // Limit a daily series to a single month (30/31 days) based on a target date string (YYYY-MM-DD)
  function restrictSurDailyToMonth(daily, endDateStr) {
    if (!daily || !Array.isArray(daily.labels) || daily.labels.length === 0) return daily;
    const monthKey = (endDateStr || '').slice(0, 7) || (daily.labels[daily.labels.length - 1] || '').slice(0, 7);
    if (!monthKey) return daily;
    const labels = []; const inflows = []; const production = []; const productionMWCI = []; const productionMWSI = []; const augment = [];
    const augmentMWCI = []; const augmentMWSI = [];
    for (let i = 0; i < daily.labels.length; i++) {
      const d = daily.labels[i] || '';
      if (d.slice(0, 7) === monthKey) {
        labels.push(d);
        inflows.push(daily.inflows[i]);
        production.push(daily.production[i]);
        productionMWCI.push(daily.productionMWCI?.[i] || 0);
        productionMWSI.push(daily.productionMWSI?.[i] || 0);
        augment.push(daily.augment[i]);
        augmentMWCI.push(daily.augmentMWCI?.[i] || 0);
        augmentMWSI.push(daily.augmentMWSI?.[i] || 0);
      }
    }
    return {
      labels,
      inflows,
      production,
      productionMWCI,
      productionMWSI,
      augment,
      augmentMWCI,
      augmentMWSI,
      hasData: labels.length > 0
    };
  }

  function getFilteredServiceUpdatesForChart() {
    let list = serviceUpdates.slice();
    const start = elements.serviceStartDateFilter?.value || '';
    const end = elements.serviceEndDateFilter?.value || '';
    const prov = elements.serviceProviderFilter?.value || '';
    const provCanon = prov ? canonProvider(prov) : '';
    if (start) list = list.filter(s => safeDateCompare((s.date || ''), start) >= 0);
    if (end) list = list.filter(s => safeDateCompare((s.date || ''), end) <= 0);
    if (provCanon) list = list.filter(s => canonProvider(s.provider || '') === provCanon);
    return list;
  }

  function renderSurChart() {
    const canvas = elements.surChartCanvas;
    const emptyEl = elements.surChartEmpty;
    if (!canvas) return;
    const list = getFilteredServiceUpdatesForChart();
    const prov = elements.serviceProviderFilter?.value || '';
    const daily = buildSurDailyAggregation(list);
    const view = elements.surChartView?.value || 'daily';
    const aggMode = 'sum'; // always totals
    let series = null;
    if (view === 'monthly') {
      const endDate = elements.serviceEndDateFilter?.value || daily.labels[daily.labels.length - 1] || elements.serviceStartDateFilter?.value || '';
      const yearStr = (endDate || '').slice(0, 4);
      series = buildSurMonthlyTotalsForYear(daily, yearStr);
    } else if (view === 'yearly') {
      series = buildSurYearlyFromDaily(daily, aggMode);
    } else {
      const endDate = elements.serviceEndDateFilter?.value || daily.labels[daily.labels.length - 1] || '';
      series = restrictSurDailyToMonth(daily, endDate);
    }
    if (!series.hasData) {
      if (emptyEl) emptyEl.style.display = 'block';
      canvas.style.display = 'none';
      if (surChart) { surChart.destroy(); surChart = null; }
      return;
    }
    if (emptyEl) emptyEl.style.display = 'none';
    canvas.style.display = 'block';
    const data = {
      labels: series.labels,
      datasets: [
        {
          type: 'bar',
          label: 'Treatment Production - MWCI',
          data: series.productionMWCI || series.production,
          backgroundColor: 'rgba(0,51,102,0.7)', // navy blue
          borderColor: '#003366',
          borderWidth: 1,
          stack: 'supply',
          yAxisID: 'y',
          hidden: prov === 'MWSI',
        },
        {
          type: 'bar',
          label: 'Treatment Production - MWSI',
          data: series.productionMWSI || series.production,
          backgroundColor: 'rgba(116,192,252,0.7)', // light blue
          borderColor: '#74C0FC',
          borderWidth: 1,
          stack: 'supply',
          yAxisID: 'y',
          hidden: prov === 'MWCI',
        },
        {
          type: 'bar',
          label: 'Supply Augmentation',
          data: series.augment,
          backgroundColor: 'rgba(255,193,7,0.6)',
          borderColor: '#ffc107',
          borderWidth: 1,
          stack: 'supply',
          yAxisID: 'y',
        },
        {
          type: 'line',
          label: 'Raw Inflows',
          data: series.inflows,
          borderColor: '#198754',
          backgroundColor: 'rgba(25,135,84,0.2)',
          tension: 0.25,
          yAxisID: 'y',
        }
      ]
    };
    const isMobile = window.matchMedia('(max-width: 576px)').matches;
    const tickFontSize = isMobile ? 10 : 12;
    const maxDailyTicks = isMobile ? 6 : 12;
    const maxMonthlyTicks = isMobile ? 6 : 12;
    const xTicks = (view === 'daily')
      ? {
        autoSkip: true,
        maxTicksLimit: maxDailyTicks,
        maxRotation: 0,
        minRotation: 0,
        callback: function (val) {
          const v = (this && this.getLabelForValue) ? this.getLabelForValue(val) : val;
          return String(v).slice(5); // show MM-DD for daily
        }
      }
      : {
        autoSkip: true,
        maxTicksLimit: maxMonthlyTicks,
        maxRotation: 0,
        minRotation: 0,
      };
    const yTitle = 'MLD (Total)';
    const opts = {
      responsive: true,
      maintainAspectRatio: true,
      aspectRatio: isMobile ? 2.2 : 3.0,
      resizeDelay: 150,
      interaction: { mode: 'index', intersect: false },
      layout: { padding: { top: 4, right: 6, bottom: 2, left: 4 } },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: tickFontSize } } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const val = fmtNum(ctx.parsed.y);
              const lines = [`${ctx.dataset.label}: ${val} MLD`];
              if (ctx.dataset.label === 'Supply Augmentation') {
                const mwci = series.augmentMWCI?.[ctx.dataIndex] || 0;
                const mwsi = series.augmentMWSI?.[ctx.dataIndex] || 0;
                lines.push(`  - MWCI: ${fmtNum(mwci)} MLD`);
                lines.push(`  - MWSI: ${fmtNum(mwsi)} MLD`);
              }
              return lines;
            },
            footer: (items) => {
              try {
                const supplySum = items
                  .filter(i => i.dataset.stack === 'supply')
                  .reduce((sum, i) => sum + (i.parsed?.y || 0), 0);
                return `Net Supply (Total): ${fmtNum(supplySum)} MLD`;
              } catch (_) { return ''; }
            }
          }
        }
      },
      elements: { bar: { borderRadius: 4, borderSkipped: 'bottom' } },
      animation: { duration: isMobile ? 200 : 400 },
      scales: {
        x: { title: { display: true, text: (view === 'monthly' ? 'Month' : (view === 'yearly' ? 'Year' : 'Date (Plants Date)')) }, stacked: true, ticks: xTicks, grid: { color: 'rgba(0,0,0,0.06)' } },
        y: { title: { display: true, text: yTitle }, beginAtZero: true, stacked: true, ticks: { font: { size: tickFontSize } }, grid: { color: 'rgba(0,0,0,0.06)' } }
      }
    };
    if (surChart) {
      surChart.data = data;
      surChart.options = opts;
      surChart.update();
    } else {
      const ctx = canvas.getContext('2d');
      surChart = new Chart(ctx, { type: 'bar', data, options: opts });
    }
  }

  // ---- Dashboard rendering ----
  function aggregateDeepwellMonthlyTotals() {
    const list = DeepwellsFeatureInstance.getDeepwells();
    const byProv = { MWCI: {}, MWSI: {} };
    (list || []).forEach(dw => {
      const prov = (dw.provider || '').toUpperCase();
      if (!(prov in byProv)) return;
      (dw.months || []).forEach(m => {
        const key = m?.month;
        const val = Number(m?.prod || 0);
        if (!key || !isFinite(val)) return;
        byProv[prov][key] = (byProv[prov][key] || 0) + val;
      });
    });
    const months = Array.from(new Set([
      ...Object.keys(byProv.MWCI),
      ...Object.keys(byProv.MWSI)
    ])).sort();
    return {
      labels: months.map(monthKey => {
        if (!monthKey || typeof monthKey !== 'string' || monthKey.length < 7) return monthKey || '';
        const [y, m] = monthKey.split('-');
        const d = new Date(Number(y), Number(m) - 1, 1);
        return isNaN(d.getTime()) ? monthKey : d.toLocaleString('en-US', { month: 'short', year: 'numeric' });
      }),
      // Convert cu.m to ML (Million Liters): 1 cu.m = 1000 liters = 0.001 ML
      mwci: months.map(k => +((byProv.MWCI[k] || 0) / 1000).toFixed(2)),
      mwsi: months.map(k => +((byProv.MWSI[k] || 0) / 1000).toFixed(2)),
      hasData: months.length > 0
    };
  }

  let dashDwMonthlyChartSignature = '';
  function renderDeepwellMonthlyChartDash() {
    const canvas = elements.dwMonthlyChartDashCanvas;
    const emptyEl = elements.dashDwChartEmpty;
    if (!canvas) return; // dashboard not present
    const agg = aggregateDeepwellMonthlyTotals();

    // Create a data signature to avoid re-rendering if data is the same
    const signature = JSON.stringify(agg.labels) + JSON.stringify(agg.mwci) + JSON.stringify(agg.mwsi);
    if (dashDwMonthlyChartSignature === signature && dashDwMonthlyChart) return;

    if (!agg.hasData) {
      if (emptyEl) emptyEl.style.display = 'block';
      canvas.style.display = 'none';
      if (dashDwMonthlyChart) { dashDwMonthlyChart.destroy(); dashDwMonthlyChart = null; }
      return;
    }
    if (emptyEl) emptyEl.style.display = 'none';
    canvas.style.display = 'block';
    const data = {
      labels: agg.labels,
      datasets: [
        {
          label: 'MWCI',
          data: agg.mwci,
          borderColor: '#0d6efd',
          backgroundColor: 'rgba(13,110,253,0.15)',
          tension: 0.25,
          fill: true,
          borderWidth: 2,
          pointRadius: 1.5,
        },
        {
          label: 'MWSI',
          data: agg.mwsi,
          borderColor: '#198754',
          backgroundColor: 'rgba(25,135,84,0.15)',
          tension: 0.25,
          fill: true,
          borderWidth: 2,
          pointRadius: 1.5,
        }
      ]
    };
    const isMobile = window.matchMedia('(max-width: 576px)').matches;
    const tickFontSize = isMobile ? 10 : 12;
    const maxMonthTicks = isMobile ? 6 : 12;
    const opts = {
      responsive: true,
      maintainAspectRatio: true,
      aspectRatio: isMobile ? 1.6 : 2.2,
      resizeDelay: 150,
      interaction: { mode: 'index', intersect: false },
      layout: { padding: { top: 4, right: 6, bottom: 2, left: 4 } },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: tickFontSize } } },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${fmtNum(ctx.parsed.y)}`,
            footer: (items) => {
              const total = items.reduce((sum, item) => sum + (item.parsed.y || 0), 0);
              return `Total Production: ${fmtNum(total)} ML`;
            }
          }
        },
      },
      animation: { duration: isMobile ? 200 : 400 },
      scales: {
        x: { title: { display: true, text: 'Month' }, ticks: { autoSkip: true, maxTicksLimit: maxMonthTicks, font: { size: tickFontSize } }, grid: { color: 'rgba(0,0,0,0.06)' } },
        y: { title: { display: true, text: 'Production (ML)' }, beginAtZero: true, ticks: { font: { size: tickFontSize } }, grid: { color: 'rgba(0,0,0,0.06)' } }
      }
    };
    if (dashDwMonthlyChart) {
      dashDwMonthlyChart.data = data;
      dashDwMonthlyChart.options = opts;
      dashDwMonthlyChart.update();
    } else {
      const ctx = canvas.getContext('2d');
      dashDwMonthlyChart = new Chart(ctx, { type: 'line', data, options: opts });
    }
    dashDwMonthlyChartSignature = signature;
  }

  let dashSurChartSignature = '';
  function renderSurChartDash() {
    const canvas = elements.surChartDashCanvas;
    const emptyEl = elements.dashSurChartEmpty;
    if (!canvas) return;
    const daily = buildSurDailyAggregation(serviceUpdates || []);
    // Monthly totals for the latest year from data
    const last = daily.labels[daily.labels.length - 1] || '';
    const yearStr = last ? last.slice(0, 4) : String(new Date().getFullYear());
    const series = buildSurMonthlyTotalsForYear(daily, yearStr);

    // Create a data signature to avoid re-rendering if data is the same
    const signature = JSON.stringify(series.labels) + JSON.stringify(series.productionMWCI) + JSON.stringify(series.productionMWSI) + JSON.stringify(series.augment) + JSON.stringify(series.inflows);
    if (dashSurChartSignature === signature && dashSurChart) return;

    if (!series.hasData) {
      if (emptyEl) emptyEl.style.display = 'block';
      canvas.style.display = 'none';
      if (dashSurChart) { dashSurChart.destroy(); dashSurChart = null; }
      return;
    }
    if (emptyEl) emptyEl.style.display = 'none';
    canvas.style.display = 'block';
    const data = {
      labels: series.labels,
      datasets: [
        {
          type: 'bar',
          label: 'Treatment Production - MWCI',
          data: series.productionMWCI || series.production,
          backgroundColor: 'rgba(0,51,102,0.7)',
          borderColor: '#003366',
          borderWidth: 1,
          stack: 'supply',
          yAxisID: 'y',
        },
        {
          type: 'bar',
          label: 'Treatment Production - MWSI',
          data: series.productionMWSI || series.production,
          backgroundColor: 'rgba(116,192,252,0.7)',
          borderColor: '#74C0FC',
          borderWidth: 1,
          stack: 'supply',
          yAxisID: 'y',
        },
        {
          type: 'bar',
          label: 'Supply Augmentation',
          data: series.augment,
          backgroundColor: 'rgba(255,193,7,0.6)',
          borderColor: '#ffc107',
          borderWidth: 1,
          stack: 'supply',
          yAxisID: 'y',
        },
        {
          type: 'line',
          label: 'Raw Inflows',
          data: series.inflows,
          borderColor: '#198754',
          backgroundColor: 'rgba(25,135,84,0.2)',
          tension: 0.25,
          yAxisID: 'y',
        }
      ]
    };
    const isMobile = window.matchMedia('(max-width: 576px)').matches;
    const tickFontSize = isMobile ? 10 : 12;
    const maxMonthTicks = isMobile ? 6 : 12;
    const opts = {
      responsive: true,
      maintainAspectRatio: false,
      devicePixelRatio: window.devicePixelRatio || 1,
      resizeDelay: 0,
      interaction: { mode: 'index', intersect: false },
      layout: { padding: { top: 4, right: 6, bottom: 2, left: 4 } },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: tickFontSize } } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const val = fmtNum(ctx.parsed.y);
              const lines = [`${ctx.dataset.label}: ${val} ML`];
              if (ctx.dataset.label === 'Supply Augmentation') {
                const mwci = series.augmentMWCI?.[ctx.dataIndex] || 0;
                const mwsi = series.augmentMWSI?.[ctx.dataIndex] || 0;
                lines.push(`  - MWCI: ${fmtNum(mwci)} ML`);
                lines.push(`  - MWSI: ${fmtNum(mwsi)} ML`);
              }
              return lines;
            },
            footer: (items) => {
              const total = items
                .filter(i => i.dataset.stack === 'supply')
                .reduce((sum, i) => sum + (i.parsed.y || 0), 0);
              return `Net Supply (Total): ${fmtNum(total)} ML`;
            }
          }
        }
      },
      animation: { duration: isMobile ? 200 : 400 },
      elements: { bar: { borderRadius: 4, borderSkipped: 'bottom' } },
      scales: {
        x: { title: { display: true, text: 'Month' }, stacked: true, ticks: { autoSkip: true, maxTicksLimit: maxMonthTicks, font: { size: tickFontSize } }, grid: { color: 'rgba(0,0,0,0.06)' } },
        y: { title: { display: true, text: 'ML' }, beginAtZero: true, stacked: true, ticks: { font: { size: tickFontSize } }, grid: { color: 'rgba(0,0,0,0.06)' } }
      }
    };
    if (dashSurChart) {
      dashSurChart.data = data;
      dashSurChart.options = opts;
      dashSurChart.resize();
      dashSurChart.update('none');
    } else {
      const ctx = canvas.getContext('2d');
      dashSurChart = new Chart(ctx, { type: 'bar', data, options: opts });
    }
    dashSurChartSignature = signature;
  }
  // ---- Deepwell Map Logic ----
  const DEEPWELL_MAP_CENTER = [14.47, 121.04];
  const DEEPWELL_MAP_ZOOM = 10;

  function parseDmsCoordinate(value) {
    if (!value) return NaN;
    if (typeof value === 'number') return value;
    const text = String(value).trim();
    const regex = /(\d+)[Â°\s]+(\d+)['\s]+(\d+(?:\.\d+)?)"?/;
    const match = text.match(regex);
    if (!match) {
      const val = parseFloat(text);
      return isFinite(val) ? val : NaN;
    }
    const deg = parseFloat(match[1]);
    const min = parseFloat(match[2]);
    const sec = parseFloat(match[3]);
    let dd = deg + min / 60 + sec / 3600;
    if (text.includes('S') || text.includes('W')) dd = -dd;
    return dd;
  }

  function getDeepwellMarkerClass(status) {
    const s = (status || '').toLowerCase();
    if (s.includes('active') && !s.includes('inactive')) return 'marker-active';
    if (s.includes('standby')) return 'marker-standby';
    if (s.includes('inactive')) return 'marker-inactive';
    return 'marker-other';
  }

  function createCustomMarker(markerClass) {
    return L.divIcon({
      className: '',
      html: `<div class="deepwell-custom-marker ${markerClass}"></div>`,
      iconSize: [20, 20],
      iconAnchor: [10, 10],
      popupAnchor: [0, -10]
    });
  }

  /** Destroy an existing Leaflet map instance and clear its container */
  function destroyMap(mapInstance, container) {
    if (mapInstance) {
      try { mapInstance.remove(); } catch (_) {}
    }
    if (container) container.innerHTML = '';
    return null;
  }

  /** Create a fresh Leaflet map in the given container */
  function createMap(container, interactive) {
    if (!container || !window.L) return null;
    // ensure container is empty
    container.innerHTML = '';
    const map = L.map(container, {
      zoomControl: interactive,
      dragging: interactive,
      scrollWheelZoom: interactive,
      doubleClickZoom: interactive,
      boxZoom: interactive,
      keyboard: interactive,
      tap: interactive,
      attributionControl: interactive
    }).setView(DEEPWELL_MAP_CENTER, DEEPWELL_MAP_ZOOM);

    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap'
    }).addTo(map);

    return map;
  }

  /** Add markers to a map and optionally fit bounds */
  function addMapMarkers(map, wells, interactive) {
    if (!map || !window.L) return;
    const bounds = [];
    wells.forEach(dw => {
      const lat = parseDmsCoordinate(dw.latitude);
      const lng = parseDmsCoordinate(dw.longitude);
      if (isNaN(lat) || isNaN(lng)) return;
      const mc = getDeepwellMarkerClass(dw.status);
      const marker = L.marker([lat, lng], { icon: createCustomMarker(mc) }).addTo(map);
      bounds.push([lat, lng]);

      // Add tooltip showing the name on hover
      marker.bindTooltip(dw.name || 'Deepwell', {
        direction: 'top',
        offset: [0, -10],
        className: 'custom-map-tooltip'
      });

      if (interactive) {
        marker.bindPopup(`
          <div style="min-width:140px">
            <div class="fw-bold text-primary mb-1">${dw.name || 'Deepwell'}</div>
            <div class="small"><b>Status:</b> ${dw.status || '-'}</div>
            <div class="small"><b>Provider:</b> ${dw.provider || '-'}</div>
            <div class="small text-muted">${dw.municipality || dw.location || ''}</div>
          </div>
        `);
      }
    });
    if (bounds.length > 0) {
      map.fitBounds(bounds, { padding: [30, 30], maxZoom: 13 });
    }
  }

  /**
   * Render the preview map on the dashboard card.
   * Destroys and recreates the map each time data changes to avoid stale tile issues.
   */
  function renderDeepwellMapDash() {
    const previewEl = elements.deepwellMapPreview;
    if (!previewEl || !window.L) return;

    const wells = DeepwellsFeatureInstance.getDeepwells();
    const signature = wells.map(w => `${w.id}:${w.status}:${w.latitude}:${w.longitude}`).join('|');

    // Update empty states
    const hasMapped = wells.some(w => !isNaN(parseDmsCoordinate(w.latitude)) && !isNaN(parseDmsCoordinate(w.longitude)));
    if (elements.dashDeepwellMapEmpty) elements.dashDeepwellMapEmpty.style.display = hasMapped ? 'none' : 'block';
    if (elements.deepwellMapModalEmpty) elements.deepwellMapModalEmpty.style.display = hasMapped ? 'none' : 'block';

    const refitBounds = () => {
      if (!dashDeepwellMapPreview) return;
      const bounds = wells.map(w => [parseDmsCoordinate(w.latitude), parseDmsCoordinate(w.longitude)]).filter(b => !isNaN(b[0]) && !isNaN(b[1]));
      if (bounds.length > 0) {
        try { dashDeepwellMapPreview.fitBounds(bounds, { padding: [15, 15], maxZoom: 13 }); } catch (_) {}
      }
    };

    // Only rebuild if data changed
    if (dashDeepwellMapPreviewSignature === signature && dashDeepwellMapPreview) {
      // Multiple invalidateSize and refitBounds passes to ensure all tiles load and map centers when container becomes visible
      setTimeout(() => { try { dashDeepwellMapPreview.invalidateSize(); refitBounds(); } catch(_){} }, 50);
      setTimeout(() => { try { dashDeepwellMapPreview.invalidateSize(); refitBounds(); } catch(_){} }, 300);
      setTimeout(() => { try { dashDeepwellMapPreview.invalidateSize(); refitBounds(); } catch(_){} }, 800);
      return;
    }

    // Destroy old map and create fresh
    dashDeepwellMapPreview = destroyMap(dashDeepwellMapPreview, previewEl);

    // Wait for the container to be painted with correct dimensions
    requestAnimationFrame(() => {
      dashDeepwellMapPreview = createMap(previewEl, false);
      if (!dashDeepwellMapPreview) return;
      addMapMarkers(dashDeepwellMapPreview, wells, false);
      dashDeepwellMapPreviewSignature = signature;
      // Multiple invalidateSize and refitBounds passes
      setTimeout(() => { try { dashDeepwellMapPreview.invalidateSize(); refitBounds(); } catch(_){} }, 50);
      setTimeout(() => { try { dashDeepwellMapPreview.invalidateSize(); refitBounds(); } catch(_){} }, 300);
      setTimeout(() => { try { dashDeepwellMapPreview.invalidateSize(); refitBounds(); } catch(_){} }, 800);
    });
  }

  /**
   * Render the detail map inside the modal.
   * Called only when the modal is fully shown (visible with correct dimensions).
   */
  function renderDeepwellMapDetail() {
    const detailEl = elements.deepwellMapDetail;
    if (!detailEl || !window.L) return;

    const wells = DeepwellsFeatureInstance.getDeepwells();

    // Destroy any existing detail map and recreate from scratch
    dashDeepwellMapDetail = destroyMap(dashDeepwellMapDetail, detailEl);

    requestAnimationFrame(() => {
      dashDeepwellMapDetail = createMap(detailEl, true);
      if (!dashDeepwellMapDetail) return;
      addMapMarkers(dashDeepwellMapDetail, wells, true);
      // Multiple invalidateSize passes
      setTimeout(() => { try { dashDeepwellMapDetail.invalidateSize(); } catch(_){} }, 50);
      setTimeout(() => { try { dashDeepwellMapDetail.invalidateSize(); } catch(_){} }, 300);
      setTimeout(() => { try { dashDeepwellMapDetail.invalidateSize(); } catch(_){} }, 800);
    });
  }


  // ---- Dashboard Presentations per month (current year) ----
  function renderPresentationsChartDash() {
    const canvas = elements.presentationChartDashCanvas;
    const emptyEl = elements.dashPresentationChartEmpty;
    if (!canvas) return; // dashboard not present
    // Count presentations for current calendar year
    const year = new Date().getFullYear();
    const totals = new Array(12).fill(0);
    (presentations || []).forEach(p => {
      const d = normalizeYmd(p.date);
      if (!d || !d.startsWith(String(year) + '-')) return;
      const m = parseInt(d.slice(5, 7), 10) - 1;
      if (m < 0 || m > 11) return;
      totals[m]++;
    });
    const hasData = totals.some(v => v > 0);

    // Create a data signature to avoid re-rendering if data is the same
    const signature = JSON.stringify(totals) + String(year);
    if (dashPresentationsChartSignature === signature && dashPresentationsChart) return;

    if (!hasData) {
      if (emptyEl) emptyEl.style.display = 'block';
      canvas.style.display = 'none';
      try { if (dashPresentationsChart) { dashPresentationsChart.destroy(); dashPresentationsChart = null; } } catch (_) {/* noop */ }

      return;
    }
    if (emptyEl) emptyEl.style.display = 'none';
    canvas.style.display = 'block';
    // Trim trailing zero months for cleaner look
    let lastIdx = -1; for (let i = totals.length - 1; i >= 0; i--) { if (totals[i] > 0) { lastIdx = i; break; } }
    const labelsFull = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const labels = labelsFull.slice(0, (lastIdx >= 0 ? lastIdx + 1 : totals.length));
    const data = totals.slice(0, (lastIdx >= 0 ? lastIdx + 1 : totals.length));
    try { if (dashPresentationsChart) { dashPresentationsChart.destroy(); } } catch (_) {/* noop */ }

    const ctx = canvas.getContext('2d');
    const grad = ctx.createLinearGradient(0, 0, 0, canvas.height);
    grad.addColorStop(0, 'rgba(13,110,253,0.25)');
    grad.addColorStop(1, 'rgba(13,110,253,0.02)');
    const yMax = Math.max(...data);
    const suggestedMax = yMax > 0 ? yMax + 1 : 5;
    dashPresentationsChart = new Chart(ctx, {

      type: 'line',
      data: {
        labels, datasets: [{
          label: `Presentations (${year})`,
          data,
          borderColor: '#0d6efd',
          backgroundColor: grad,
          pointBackgroundColor: '#0d6efd',
          pointBorderColor: '#0d6efd',
          pointRadius: 3,
          borderWidth: 2,
          tension: 0.3,
          fill: true
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false } },
        scales: {
          y: { beginAtZero: true, suggestedMax, ticks: { stepSize: 1, color: '#64748b', font: { size: 10 } }, grid: { borderDash: [3, 3], drawBorder: false, color: 'rgba(0,0,0,0.05)' } },
          x: { ticks: { color: '#64748b', font: { size: 10 } }, grid: { display: false } }
        }
      }
    });
    dashPresentationsChartSignature = signature;
  }

  function renderDashboard() {
    // Update metric cards
    const projectsList = ProjectsFeatureInstance.getProjects();
    const deepwellsList = DeepwellsFeatureInstance.getDeepwells();
    if (elements.dashProjectsTotal) elements.dashProjectsTotal.textContent = fmtNum(projectsList.length);
    if (elements.dashDeepwellsTotal) elements.dashDeepwellsTotal.textContent = fmtNum(deepwellsList.length);
    // Projects breakdown: Ongoing (includes Delayed) and Completed
    try {
      const pStatuses = projectsList.map(p => (typeof ProjectsFeatureInstance.getProjectStatus === 'function') ? ProjectsFeatureInstance.getProjectStatus(p) : (p.status || ''));
      const ongoing = pStatuses.filter(s => s === 'On-going' || s === 'Delayed').length;
      const completed = pStatuses.filter(s => s === 'Completed').length;
      if (elements.dashProjectsOngoing) elements.dashProjectsOngoing.textContent = fmtNum(ongoing);
      if (elements.dashProjectsCompleted) elements.dashProjectsCompleted.textContent = fmtNum(completed);
      // By Implementing Agency (MWCI/MWSI/MWSS)
      const agencyNorm = a => (a || '').toString().trim().toUpperCase();
      const cntMWCI = projectsList.filter(p => agencyNorm(p.implementingAgency) === 'MWCI').length;
      const cntMWSI = projectsList.filter(p => agencyNorm(p.implementingAgency) === 'MWSI').length;
      const cntMWSS = projectsList.filter(p => agencyNorm(p.implementingAgency) === 'MWSS').length;
      if (elements.dashProjectsMWCI) elements.dashProjectsMWCI.textContent = fmtNum(cntMWCI);
      if (elements.dashProjectsMWSI) elements.dashProjectsMWSI.textContent = fmtNum(cntMWSI);
      if (elements.dashProjectsMWSS) elements.dashProjectsMWSS.textContent = fmtNum(cntMWSS);

      // Projects completed recently (current year) based on latest accomplishment date
      const currentYear = new Date().getFullYear();
      const completedRecent = projectsList.filter(p => {
        const s = (typeof ProjectsFeatureInstance.getProjectStatus === 'function') ? ProjectsFeatureInstance.getProjectStatus(p) : (p.status || '');
        if (s !== 'Completed') return false;
        const acc = p.accomplishments || [];
        if (acc.length === 0) return false; // has status completed but no record? treat as old/unknown
        // get latest accomplishment logic same as projects.js
        const latest = acc.slice().sort((a, b) => new Date(b.date) - new Date(a.date))[0];
        if (!latest || !latest.date) return false;
        return new Date(latest.date).getFullYear() === currentYear;
      }).length;
      if (elements.dashProjectsCompletedRecently) {
        elements.dashProjectsCompletedRecently.textContent = '+' + fmtNum(completedRecent);
      }
    } catch (_) {/* noop */ }

    // Deepwell breakdown by status (Active/Inactive)
    try {
      const latestMonths = DeepwellsFeatureInstance.getLatestMonthsByProvider();
      let dwActive = 0;
      let dwInactive = 0;
      let dwStandby = 0;

      const norm = s => (s || '').toString().trim().toLowerCase();
      const canon = s => norm(s).replace(/[^a-z0-9]/g, ''); // normalize for variant matching
      const standbyCodes = new Set(['rto', 'rfo', 'ud']);

      deepwellsList.forEach(dw => {
        const prov = (dw.provider || '').toUpperCase();
        const refMonth = latestMonths[prov] || latestMonths.MWCI;
        
        // On-the-fly "Active" check based on latest data for this provider
        const monthSet = new Set((dw.months || []).map(m => String(m.month || '').trim()));
        const last2 = DeepwellsFeatureInstance.generateMonthKeys(2, refMonth);
        const hasRecentData = last2.some(k => monthSet.has(k));

        if (hasRecentData) {
          dwActive++;
        } else {
          // Fallback to existing status categorization for non-active wells
          const st = norm(dw.status);
          const c = canon(dw.status);
          
          if (st === 'inactive') {
            dwInactive++;
          } else if (standbyCodes.has(c) || st.includes('standby') || c.includes('standby')) {
            dwStandby++;
          } else if (st.includes('ready') && (st.includes('operate') || st.includes('operation') || st.includes('ops'))) {
            dwStandby++;
          } else if ((st.includes('under') || st.includes('u/')) && (st.includes('dev') || st.includes('devel') || st.includes('devt'))) {
            dwStandby++;
          } else {
            // Check if it should be auto-standby/inactive per logic if it's not active
            const auto = DeepwellsFeatureInstance.computeDeepwellAutoStatus(dw.months, dw.status, refMonth);
            const as = (auto.status || '').toLowerCase();
            if (as === 'standby') dwStandby++;
            else if (as === 'inactive') dwInactive++;
          }
        }
      });

      if (elements.dashDeepwellsActive) elements.dashDeepwellsActive.textContent = fmtNum(dwActive);
      if (elements.dashDeepwellsInactive) elements.dashDeepwellsInactive.textContent = fmtNum(dwInactive);
      if (elements.dashDeepwellsStandby) elements.dashDeepwellsStandby.textContent = fmtNum(dwStandby);
      
      const dwOther = Math.max(0, deepwellsList.length - (dwActive + dwInactive + dwStandby));
      if (elements.dashDeepwellsOther) elements.dashDeepwellsOther.textContent = fmtNum(dwOther);
    } catch (err) { console.error('[Dashboard] Deepwell stats error:', err); }

    // Deepwell YTD Production by provider (MWCI/MWSI) in MLD
    try {
      const currentYear = new Date().getFullYear().toString();
      let mwciYtd = 0, mwsiYtd = 0;
      (deepwellsList || []).forEach(dw => {
        const prov = (dw.provider || '').toUpperCase();
        (dw.months || []).forEach(m => {
          const monthKey = m?.month || ''; // e.g., "2025-01"
          if (!monthKey.startsWith(currentYear)) return;
          const prodCuM = Number(m?.prod || 0);
          if (!isFinite(prodCuM)) return;
          // Convert cubic meters to MLD (divide by 1000)
          const mld = prodCuM / 1000;
          if (prov === 'MWCI') mwciYtd += mld;
          else if (prov === 'MWSI') mwsiYtd += mld;
        });
      });
      if (elements.dashDwYtdYear) elements.dashDwYtdYear.textContent = `(${currentYear})`;
      if (elements.dashDwYtdMwci) elements.dashDwYtdMwci.textContent = mwciYtd > 0 ? `${fmtNum(mwciYtd.toFixed(2))} ML` : '- ML';
      if (elements.dashDwYtdMwsi) elements.dashDwYtdMwsi.textContent = mwsiYtd > 0 ? `${fmtNum(mwsiYtd.toFixed(2))} ML` : '- ML';
    } catch (_) {/* noop */ }

    // Today's dam elevations (or latest available) from Service Updates
    try {
      const toNum = v => {
        if (v === null || v === undefined) return NaN;
        const n = parseFloat(v.toString().replace(/,/g, ''));
        return Number.isFinite(n) ? n : NaN;
      };
      const toDateObj = v => {
        if (!v) return null;
        try {
          if (typeof v?.toDate === 'function') return v.toDate();
          if (typeof v === 'object' && typeof v.seconds === 'number') return new Date(v.seconds * 1000);
          const d = new Date(v);
          return isFinite(d.getTime()) ? d : null;
        } catch (_) { return null; }
      };
      const ymd = d => {
        if (!d) return '';
        const yy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yy}-${mm}-${dd}`;
      };
      const list = (serviceUpdates || []).slice().filter(s => (s.angat != null || s.ipo != null || s.laMesa != null));
      try { console.debug('[Dashboard] serviceUpdates:', (serviceUpdates || []).length, 'filtered (has dam any):', list.length); } catch (_) { }
      if (list.length) {
        const sorted = list.slice().sort((a, b) => {
          const bd = toDateObj(b.date || b.timestamp || b.createdAt);
          const ad = toDateObj(a.date || a.timestamp || a.createdAt);
          return (bd ? bd.getTime() : 0) - (ad ? ad.getTime() : 0);
        });
        const latest = sorted.find(s => {
          const a = toNum(s.angat), i = toNum(s.ipo), l = toNum(s.laMesa);
          return Number.isFinite(a) || Number.isFinite(i) || Number.isFinite(l);
        }) || sorted[0] || {};
        try { console.debug('[Dashboard] Latest dam doc:', { date: latest?.date || latest?.timestamp || latest?.createdAt, damAsOf: latest?.damAsOf, angat: latest?.angat, ipo: latest?.ipo, laMesa: latest?.laMesa }); } catch (_) { }
        const setDam = (el, v) => {
          if (!el) return;
          const n = toNum(v);
          if (Number.isFinite(n)) el.textContent = fmt2(n);
          else if (v != null && v !== '') el.textContent = v;
          else el.textContent = '-';
        };
        setDam(elements.dashAngat, latest.angat);
        setDam(elements.dashIpo, latest.ipo);
        setDam(elements.dashLaMesa, latest.laMesa);
        if (elements.dashDamAsOf) {
          const d = toDateObj(latest.date || latest.timestamp || latest.createdAt);
          // Show both date and time if available, else gracefully fall back
          if (latest.damAsOf && d) {
            elements.dashDamAsOf.textContent = `as of ${ymd(d)} ${latest.damAsOf}`;
          } else if (latest.damAsOf) {
            elements.dashDamAsOf.textContent = `as of ${latest.damAsOf}`;
          } else if (d) {
            elements.dashDamAsOf.textContent = `as of ${ymd(d)}`;
          } else {
            elements.dashDamAsOf.textContent = '';
          }
        }
      } else {
        if (elements.dashAngat) elements.dashAngat.textContent = '-';
        if (elements.dashIpo) elements.dashIpo.textContent = '-';
        if (elements.dashLaMesa) elements.dashLaMesa.textContent = '-';
        if (elements.dashDamAsOf) elements.dashDamAsOf.textContent = '';
      }
    } catch (_) {/* noop */ }

    // YTD High/Low for each dam (prefer current year; fallback to latest data year)
    try {
      const toNum = v => {
        if (v === null || v === undefined) return NaN;
        const n = parseFloat(v.toString().replace(/,/g, ''));
        return Number.isFinite(n) ? n : NaN;
      };
      const toDateObj = v => {
        if (!v) return null;
        try {
          if (typeof v?.toDate === 'function') return v.toDate();
          if (typeof v === 'object' && typeof v.seconds === 'number') return new Date(v.seconds * 1000);
          const d = new Date(v);
          return isFinite(d.getTime()) ? d : null;
        } catch (_) { return null; }
      };
      const getYear = (v) => {
        const d = toDateObj(v);
        return d ? d.getFullYear() : NaN;
      };
      const ymd = d => {
        if (!d) return '';
        const yy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yy}-${mm}-${dd}`;
      };
      // Build a list with valid dam measurements and dates
      const damList = (serviceUpdates || []).filter(s => {
        const a = toNum(s.angat); const i = toNum(s.ipo); const l = toNum(s.laMesa);
        return (Number.isFinite(a) || Number.isFinite(i) || Number.isFinite(l)) && (s.date || s.timestamp || s.createdAt);
      });
      let targetYear = new Date().getFullYear();
      const hasCurrent = damList.some(s => getYear(s.date || s.timestamp || s.createdAt) === targetYear);
      if (!hasCurrent && damList.length) {
        damList.sort((a, b) => {
          const bd = toDateObj(b.date || b.timestamp || b.createdAt);
          const ad = toDateObj(a.date || a.timestamp || a.createdAt);
          return (bd ? bd.getTime() : 0) - (ad ? ad.getTime() : 0);
        });
        targetYear = getYear(damList[0].date || damList[0].timestamp || damList[0].createdAt);
      }
      const inYear = (dstr) => getYear(dstr) === targetYear;
      const computeStat = (field) => {
        let hiVal = -Infinity, hiDate = '';
        let loVal = +Infinity, loDate = '';
        damList.forEach(s => {
          const raw = toNum(s[field]);
          const dateAny = (s.date || s.timestamp || s.createdAt || '');
          if (!Number.isFinite(raw)) return;
          if (!inYear(dateAny)) return;
          const d = toDateObj(dateAny);
          const dStr = ymd(d);
          if (raw > hiVal) { hiVal = raw; hiDate = dStr; }
          if (raw < loVal) { loVal = raw; loDate = dStr; }
        });
        if (hiVal === -Infinity || loVal === +Infinity) return null;
        return { hiVal, hiDate, loVal, loDate };
      };
      const setYtd = (prefix, stat) => {
        const hiValEl = elements[`${prefix}HiVal`];
        const hiDateEl = elements[`${prefix}HiDate`];
        const loValEl = elements[`${prefix}LoVal`];
        const loDateEl = elements[`${prefix}LoDate`];
        if (hiValEl || hiDateEl || loValEl || loDateEl) {
          // New detailed layout present
          if (hiValEl) hiValEl.textContent = stat ? fmt2(stat.hiVal) : '-';
          if (hiDateEl) hiDateEl.textContent = stat ? stat.hiDate : '';
          if (loValEl) loValEl.textContent = stat ? fmt2(stat.loVal) : '-';
          if (loDateEl) loDateEl.textContent = stat ? stat.loDate : '';
        } else {
          // Fallback to old single-line field if exists
          const lineEl = elements[prefix.replace('Ytd', 'Ytd')];
          if (lineEl) {
            if (!stat) lineEl.textContent = '-';
            else lineEl.textContent = `High ${fmt2(stat.hiVal)} m (${stat.hiDate}) - Low ${fmt2(stat.loVal)} m (${stat.loDate})`;
          }
        }
      };

      setYtd('dashAngatYtd', computeStat('angat'));
      setYtd('dashIpoYtd', computeStat('ipo'));
      setYtd('dashLaMesaYtd', computeStat('laMesa'));
      try { console.debug('[Dashboard] YTD targetYear:', targetYear, '| list size:', damList.length); } catch (_) { }
    } catch (_) {/* noop */ }
    // Render dashboard charts
    try { renderSurChartDash(); } catch (_) {/*noop*/ }
    try { renderDeepwellMonthlyChartDash(); } catch (_) {/*noop*/ }
    try { renderDeepwellMapDash(); } catch (_) {/*noop*/ }

    try { populateDamYearSelect(); } catch (_) {/*noop*/ }
    try { renderDamElevationsChartDash(); } catch (_) {/*noop*/ }
    try {
      const sel = document.getElementById('damYearSelect');
      if (sel && !sel.__damWired) { sel.addEventListener('change', () => renderDamElevationsChartDash()); sel.__damWired = true; }
    } catch (_) {/*noop*/ }
    try { populateTreatmentPlantYearSelect('MWCI'); } catch (_) {/*noop*/ }
    try { populateTreatmentPlantYearSelect('MWSI'); } catch (_) {/*noop*/ }
    try { renderTreatmentPlantProductionChartDash('MWCI'); } catch (_) {/*noop*/ }
    try { renderTreatmentPlantProductionChartDash('MWSI'); } catch (_) {/*noop*/ }
    try {
      const mwciSel = document.getElementById('treatmentPlantsMwciYearSelect');
      if (mwciSel && !mwciSel.__treatmentPlantsWired) {
        mwciSel.addEventListener('change', () => renderTreatmentPlantProductionChartDash('MWCI'));
        mwciSel.__treatmentPlantsWired = true;
      }
      const mwsiSel = document.getElementById('treatmentPlantsMwsiYearSelect');
      if (mwsiSel && !mwsiSel.__treatmentPlantsWired) {
        mwsiSel.addEventListener('change', () => renderTreatmentPlantProductionChartDash('MWSI'));
        mwsiSel.__treatmentPlantsWired = true;
      }
      wireTreatmentPlantExpandedCards();
      wireTreatmentPlantExpandedExportActions();
    } catch (_) {/*noop*/ }
  }

  // Load config for signups
  async function loadConfig() {
    try {
      const doc = await db.collection('config').doc('app').get();
      if (doc.exists) {
        const data = doc.data();
        // signup enable/disable config deprecated
      }
    } catch (err) { console.warn('Failed to load config', err); }
    updateAdminUI();
    subscribeDeepwells();
  }

  function updateAdminUI() {
    if (isAdmin) {
      // Only real admins can manage users
      elements.manageUsersBtn.style.display = 'inline-block';
      loadPendingUsers();
      loadApprovedUsers();
    } else {
      elements.manageUsersBtn.style.display = 'none';
      pendingUsers = [];
      updatePendingBadge();
    }
  }

  // Viewer-mode UI: hide Personal Tasks and all add/import/template controls
  function updateViewOnlyUI() {
    const hide = (el) => { if (!el) return; el.hidden = true; el.style.display = 'none'; el.setAttribute('aria-hidden', 'true'); try { el.disabled = true; } catch (_) { } };
    const show = (el) => { if (!el) return; el.hidden = false; el.style.display = ''; el.removeAttribute('aria-hidden'); try { el.disabled = false; } catch (_) { } };
    if (isViewOnly) {
      // Personal Tasks hidden
      hide(elements.personalTasksBtn);
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';
      hide(elements.ptmAddBtn);
      // Construction Projects add button (if present)
      const addProjectBtnEl = document.getElementById('addProjectBtn'); hide(addProjectBtnEl);
      // Reforestation
      hide(elements.addReforestationBtn);
      // Deepwells
      hide(elements.addDeepwellBtn);
      hide(elements.addDwMonthBtn);
      // Service Updates
      hide(elements.addServiceUpdateBtn);
      // Service Update (SUR): hide bulk/template/export
      hide(elements.surBulkTemplateBtn);
      hide(elements.surBulkImportBtn);
      hide(elements.surExportBtn);
      try {
        const surWrap = elements.surBulkTemplateBtn ? elements.surBulkTemplateBtn.closest('.d-flex') : null;
        if (surWrap) surWrap.style.display = 'none';
        const surCol = elements.surBulkTemplateBtn ? elements.surBulkTemplateBtn.closest('.col-md-3') : null;
        if (surCol) surCol.style.display = 'none';
      } catch (_) {/* noop */ }
    } else {
      // Non-viewer (signed-in): restore visibility; detailed access handled elsewhere
      const addProjectBtnEl = document.getElementById('addProjectBtn'); show(addProjectBtnEl);
      show(elements.personalTasksBtn);
      show(elements.addReforestationBtn);
      show(elements.addDeepwellBtn);
      show(elements.addDwMonthBtn);
      show(elements.addServiceUpdateBtn);
      show(elements.surBulkTemplateBtn);
      show(elements.surBulkImportBtn);
      show(elements.surExportBtn);
      try {
        const surWrap = elements.surBulkTemplateBtn ? elements.surBulkTemplateBtn.closest('.d-flex') : null;
        if (surWrap) surWrap.style.display = '';
        const surCol = elements.surBulkTemplateBtn ? elements.surBulkTemplateBtn.closest('.col-md-3') : null;
        if (surCol) surCol.style.display = '';
      } catch (_) {/* noop */ }
    }
  }

  async function loadPendingUsers() {
    try {
      const snap = await userService.getPending();
      pendingUsers = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      renderPendingUsersTable();
      updatePendingBadge();
    } catch (err) { console.error('Failed to load pending users', err); }
  }

  // Load and render approved users list
  async function loadApprovedUsers() {
    try {
      const snap = await userService.getApproved();
      approvedUsers = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      renderApprovedUsersTable();
      if (window.__refreshChatRecipients) window.__refreshChatRecipients();
    } catch (err) { console.error('Failed to load approved users', err); }
  }

  function renderApprovedUsersTable() {
    const tbody = elements.approvedUsersTable.querySelector('tbody');
    tbody.innerHTML = '';
    // Ensure Admin is always listed at the top, even if not present in Firestore approved users
    try {
      const present = approvedUsers.some(u => (u.email || '').trim().toLowerCase() === ADMIN_EMAIL_LOWER);
      if (!present) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${ADMIN_FULL_NAME}</td>` +
          `<td>${ADMIN_EMAIL}</td>` +
          `<td>${ADMIN_DESIGNATION}</td>` +
          `<td>${ADMIN_DEPARTMENT}</td>` +
          `<td>-</td>` +
          `<td><span class="badge text-bg-dark">Admin</span></td>` +
          `<td class="text-nowrap">-</td>`;
        tbody.appendChild(tr);
      }
    } catch (_) { }
    approvedUsers.forEach(u => {
      const tr = document.createElement('tr');
      const approvedWhen = u.updatedAt?.toDate ? u.updatedAt.toDate().toLocaleString() : (u.createdAt?.toDate ? u.createdAt.toDate().toLocaleString() : '');
      const emailLower = (u.email || '').trim().toLowerCase();
      if (emailLower === ADMIN_EMAIL_LOWER) {
        // Render Admin as a fixed row with 'Admin' access label and no actions
        tr.innerHTML = `<td>${ADMIN_FULL_NAME || u.fullName || ''}</td>` +
          `<td>${ADMIN_EMAIL}</td>` +
          `<td>${ADMIN_DESIGNATION || u.designation || ''}</td>` +
          `<td>${ADMIN_DEPARTMENT || u.department || ''}</td>` +
          `<td>${approvedWhen}</td>` +
          `<td><span class="badge text-bg-dark">Admin</span></td>` +
          `<td class="text-nowrap">-</td>`;
        tbody.appendChild(tr);
        return;
      }
      // Add placeholders for Access Level and Action columns
      tr.innerHTML = `<td>${u.fullName || ''}</td><td>${u.email || ''}</td><td>${u.designation || ''}</td><td>${u.department || ''}</td><td>${approvedWhen}</td><td></td><td class="text-nowrap"></td>`;
      const actionCell = tr.lastElementChild;
      const levelCell = actionCell.previousElementSibling;
      const select = document.createElement('select');
      select.className = 'form-select form-select-sm';
      [1, 2].forEach(l => {
        const opt = document.createElement('option');
        opt.value = l;
        opt.textContent = 'Level ' + l;
        if ((u.accessLevel || 1) === l) opt.selected = true;
        select.appendChild(opt);
      });
      select.onchange = () => updateUserAccessLevel(u, parseInt(select.value, 10));
      levelCell.appendChild(select);
      // Actions: Revoke (set approved:false) and Delete (remove record)
      const revokeBtn = document.createElement('button');
      revokeBtn.className = 'btn btn-sm btn-warning me-1';
      revokeBtn.textContent = 'Revoke';
      revokeBtn.title = 'Move back to Pending (approved=false)';
      revokeBtn.onclick = () => revokeUser(u);

      const delBtn = document.createElement('button');
      delBtn.className = 'btn btn-sm btn-danger';
      delBtn.textContent = 'Delete';
      delBtn.title = 'Delete user account';
      delBtn.onclick = () => deleteApprovedUser(u);

      actionCell.append(revokeBtn, delBtn);
      tbody.appendChild(tr);
    });
  }

  async function updateUserAccessLevel(u, level) {
    try {
      await userService.setAccessLevel(u.id, level);
      u.accessLevel = level;
    } catch (err) { alert('Failed to update access level: ' + err.message); }
  }

  // Revoke an approved user: set approved=false and refresh tables
  async function revokeUser(u) {
    try {
      await userService.setApproved(u.id, false, { updatedAt: serverTimestamp() });
      // Move user to Pending list UI
      await loadApprovedUsers();
      await loadPendingUsers();
    } catch (err) { alert('Failed to revoke user: ' + err.message); }
  }

  // Permanently delete a user document
  async function deleteApprovedUser(u) {
    try {
      const ok = confirm(`Delete the account for ${u.fullName || u.email || 'this user'}? This cannot be undone.`);
      if (!ok) return;
      await userService.deleteUser(u.id);
      await loadApprovedUsers();
    } catch (err) { alert('Failed to delete user: ' + err.message); }
  }

  function renderPendingUsersTable() {
    const tbody = elements.pendingUsersTable.querySelector('tbody');
    tbody.innerHTML = '';
    pendingUsers.forEach(u => {
      const tr = document.createElement('tr');
      const createdWhen = u.createdAt?.toDate ? u.createdAt.toDate().toLocaleString() : '';
      tr.innerHTML = `<td>${u.fullName || ''}</td><td>${u.email || ''}</td><td>${u.designation || ''}</td><td>${u.department || ''}</td><td>${createdWhen}</td><td></td>`;
      const actions = tr.lastElementChild;
      const approve = document.createElement('button');
      approve.className = 'btn btn-sm btn-success me-1';
      approve.textContent = 'Approve';
      approve.onclick = () => approveUser(u);
      const reject = document.createElement('button');
      reject.className = 'btn btn-sm btn-danger';
      reject.textContent = 'Reject';
      reject.onclick = () => rejectUser(u);
      actions.append(approve, reject);
      tbody.appendChild(tr);
    });
  }

  async function approveUser(u) {
    try {
      await userService.setApproved(u.id, true, { accessLevel: 1, updatedAt: serverTimestamp() });

      // Queue "Account Approved" email
      emailService.collection().add({
        to: [u.email],
        message: {
          subject: 'Account Approved - Construction Dashboard',
          html: `<p>Dear ${u.fullName || 'User'},</p>
                 <p>Good news! Your registration has been <strong>approved</strong>.</p>
                 <p>You can now login to the system using your registered email: <strong>${u.email}</strong></p>
                 <p>Login here: <a href="http://172.16.9.15/construction/">http://172.16.9.15/construction/</a></p>
                 <p><em>Note: You must be connected to <strong>MWSS_NET</strong> to access the system.</em></p>
                 <p>Best regards,<br>System Administrator</p>`
        }
      }).catch(err => console.error('[Approve] Email queue error:', err));

      auth.sendPasswordResetEmail(u.email).catch(() => { });
      loadPendingUsers();
      loadApprovedUsers();
    } catch (e) { alert(e.message); }
  }

  async function rejectUser(u) {
    try {
      await userService.deleteUser(u.id);
      loadPendingUsers();
    } catch (e) { alert(e.message); }
  }


  // ---- Projects (delegated to ProjectsFeature) ----
  function renderProjects() {
    ProjectsFeatureInstance.renderProjects();
  }

  // --- Photo Viewer helpers ---
  const viewerState = { list: [], index: 0, open: false };

  function showViewerImage() {
    const imgEl = document.getElementById('photoViewerImg');
    if (imgEl) imgEl.src = viewerState.list[viewerState.index] || '';
  }
  function updateViewerNav() {
    const prevBtn = document.getElementById('photoPrev');
    const nextBtn = document.getElementById('photoNext');
    const multi = viewerState.list.length > 1;
    if (prevBtn) prevBtn.style.display = multi ? 'flex' : 'none';
    if (nextBtn) nextBtn.style.display = multi ? 'flex' : 'none';
  }
  function prevPhoto() {
    if (viewerState.list.length <= 1) return;
    viewerState.index = (viewerState.index - 1 + viewerState.list.length) % viewerState.list.length;
    showViewerImage();
  }
  function nextPhoto() {
    if (viewerState.list.length <= 1) return;
    viewerState.index = (viewerState.index + 1) % viewerState.list.length;
    showViewerImage();
  }

  window.openPhotoViewer = function (indexOrSrc, list) {
    try {
      const modalEl = document.getElementById('photoViewer');
      if (!modalEl) return;
      if (Array.isArray(list)) {
        viewerState.list = list.slice();
        const n = parseInt(indexOrSrc);
        viewerState.index = isNaN(n) ? 0 : Math.max(0, Math.min(n, viewerState.list.length - 1));
      } else {
        const src = typeof indexOrSrc === 'string' ? indexOrSrc : '';
        viewerState.list = [src];
        viewerState.index = 0;
      }
      showViewerImage();
      updateViewerNav();
      bootstrap.Modal.getOrCreateInstance(modalEl).show();
    } catch (e) { console.error('openPhotoViewer error', e); }
  };

  // Setup controls and keyboard handling
  document.addEventListener('DOMContentLoaded', () => {
    const modalEl = document.getElementById('photoViewer');
    const imgEl = document.getElementById('photoViewerImg');
    const prevBtn = document.getElementById('photoPrev');
    const nextBtn = document.getElementById('photoNext');
    if (!modalEl || !imgEl) return;
    // Click image to close
    imgEl.addEventListener('click', () => bootstrap.Modal.getOrCreateInstance(modalEl).hide());
    // Prev/Next buttons
    if (prevBtn) prevBtn.addEventListener('click', e => { e.stopPropagation(); prevPhoto(); });
    if (nextBtn) nextBtn.addEventListener('click', e => { e.stopPropagation(); nextPhoto(); });

    // Keyboard
    const onKeyDown = (e) => {
      if (!viewerState.open) return;
      if (e.key === 'ArrowLeft') { e.preventDefault(); prevPhoto(); }
      else if (e.key === 'ArrowRight') { e.preventDefault(); nextPhoto(); }
    };

    modalEl.addEventListener('shown.bs.modal', () => {
      viewerState.open = true;
      document.addEventListener('keydown', onKeyDown);
    });
    modalEl.addEventListener('hidden.bs.modal', () => {
      viewerState.open = false;
      document.removeEventListener('keydown', onKeyDown);
      if (imgEl) imgEl.src = '';
      viewerState.list = [];
      viewerState.index = 0;
    });
  });

  // Ensure photo-grid 'single' class is applied when there are 1 or 2 photos
  document.addEventListener('DOMContentLoaded', () => {
    const detailsModalEl = document.getElementById('detailsModal');
    if (!detailsModalEl) return;
    detailsModalEl.addEventListener('shown.bs.modal', () => {
      try {
        const grid = detailsModalEl.querySelector('.photo-grid');
        if (grid) {
          const count = grid.querySelectorAll('.photo-tile').length;
          if (count <= 2) grid.classList.add('single'); else grid.classList.remove('single');
        }
      } catch (_) {/* noop */ }
    });
  });

  // ----- UI Cleanup Helpers -----
  function cleanDeepwellDuplicates() {
    const dwSec = document.getElementById('deepwellsSection');
    if (!dwSec) return;
    // Remove duplicate search/filter controls (keep first occurrence)
    const searchInputs = dwSec.querySelectorAll('#dwSearchInput');
    searchInputs.forEach((el, idx) => { if (idx > 0) el.closest('.col-md-3, .col-md-2, .col-md-3.text-md-end')?.remove(); });
    const providerFilters = dwSec.querySelectorAll('#dwProviderFilter');
    providerFilters.forEach((el, idx) => { if (idx > 0) el.closest('.col-md-2')?.remove(); });
    const statusFilters = dwSec.querySelectorAll('#dwStatusFilter');
    statusFilters.forEach((el, idx) => { if (idx > 0) el.closest('.col-md-2')?.remove(); });
    const prodSorts = dwSec.querySelectorAll('#dwProdSort');
    prodSorts.forEach((el, idx) => { if (idx > 0) el.closest('.col-md-2')?.remove(); });
    const addBtns = dwSec.querySelectorAll('#addDeepwellBtn');
    addBtns.forEach((btn, idx) => { if (idx > 0) btn.closest('.col-md-3')?.remove(); });
    // Remove duplicate tables (keep first)
    const tables = dwSec.querySelectorAll('#deepwellsTable');
    tables.forEach((tbl, idx) => { if (idx > 0) tbl.closest('.table-responsive')?.remove(); });
  }

  document.addEventListener('DOMContentLoaded', cleanDeepwellDuplicates);


  /* ----------------- Facebook Messenger Feature ----------------- */
  (function () {
    const messengerToggle = document.getElementById('messengerToggle');
    const messengerSidebar = document.getElementById('messengerSidebar');
    const messengerClose = document.getElementById('messengerClose');
    const messengerBadge = document.getElementById('messengerBadge');
    const msgContainer = document.getElementById('chatMessages');
    const msgInput = document.getElementById('chatInput');
    const msgSend = document.getElementById('chatSend');
    const contactsList = document.getElementById('contactsList');
    const messengerClear = document.getElementById('messengerClear');
    const chatLoadOlderBtn = document.getElementById('chatLoadOlderBtn');
    const chatLoadOlderSpinner = document.getElementById('chatLoadOlderSpinner');

    let currentChat = 'all';
    let currentChatTitle = 'General Chat';
    let currentChatAvatar = 'fa-users';

    // updateBadge is defined later in this scope (single consolidated implementation)
    // Update General Chat badge in the contacts header
    function updateGeneralBadge() {
      const badge = document.querySelector('.contact-item[data-chat="all"] .contact-badge');
      if (!badge) return;
      const count = unreadCounts['all'] || 0;
      badge.textContent = count;
      badge.style.display = count > 0 ? 'inline-flex' : 'none';
    }
    // Toast container for notifications (kept for future use)
    const toastContainer = window.UIShell?.ensureToastContainer ? window.UIShell.ensureToastContainer() : null;

    function notifyPm(fromEmail) {
      // Toast pop-up disabled at user's request; we still update the unread badge
      if (!fromEmail) return;
      updateBadge();
    }

    // Messenger UI Functions
    function showMessenger() {
      messengerSidebar.style.display = 'flex';
      document.body.style.overflow = 'hidden';
      // Clear unread for the currently open chat and refresh view
      try { switchChat(currentChat, currentChatTitle, currentChatAvatar); } catch (_) { }
    }

    function hideMessenger() {
      messengerSidebar.style.display = 'none';
      document.body.style.overflow = '';
    }

    function switchChat(chatId, title, avatar) {
      currentChat = chatId;
      currentChatTitle = title;
      currentChatAvatar = avatar;

      // Update active contact
      document.querySelectorAll('.contact-item').forEach(item => {
        item.classList.toggle('active', item.dataset.chat === chatId);
      });

      // Update chat header
      const chatTitle = document.querySelector('.chat-title');
      const chatAvatar = document.querySelector('.chat-avatar i');
      if (chatTitle) chatTitle.textContent = title;
      if (chatAvatar) chatAvatar.className = `fa ${avatar}`;

      // Clear unread count for this chat (both UID and possible email key)
      if (unreadCounts[chatId]) unreadCounts[chatId] = 0;
      if (chatId === ADMIN_EMAIL_LOWER && ADMIN_UID && unreadCounts[ADMIN_UID]) unreadCounts[ADMIN_UID] = 0;
      if (chatId === 'all') unreadCounts['all'] = 0;
      updateBadge();
      updateContactsList();

      renderMessages();
    }
    // Old chat elements removed - using Facebook Messenger elements instead
    function ensureClearBtn() {
      const user = auth.currentUser;
      const email = user && user.email ? user.email.trim().toLowerCase() : '';
      const isAdminNow = email === ADMIN_EMAIL_LOWER;
      if (!isAdminNow) return;

      // Show admin clear button in messenger
      if (messengerClear) {
        messengerClear.classList.remove('d-none');
      }
    }
    function attachClearHandler() {
      // Clear handler moved to Facebook Messenger implementation
      // No longer needed here
    }
    ensureClearBtn();
    if (!messengerToggle) return; // html not loaded yet

    // Messenger toggle is now fixed in the header - no drag functionality needed
    // Clear any old position data that might affect the button
    try { localStorage.removeItem('messengerTogglePos'); } catch (_) { }

    auth.onAuthStateChanged(() => {
      ensureClearBtn();
    });


    let msgsUnsub = null;
    const CHAT_PAGE_SIZE = 50; // default chat window size
    let msgsCursor = null;     // cursor to page older messages (last doc of current window)
    let isLoadingOlder = false;
    let suppressAutoScroll = false; // avoid jumping to bottom when prepending older
    const unreadCounts = {}; // {fromId: number}
    let ADMIN_UID = null;
    let allMessages = [];
    // Update main toggle badge and general chat badge
    function updateBadge() {
      // Sum all unread counts
      let total = 0;
      for (const k in unreadCounts) { total += Number(unreadCounts[k] || 0); }
      if (messengerBadge) {
        if (total > 0) {
          messengerBadge.textContent = String(total);
          messengerBadge.style.display = 'inline-block';
        } else {
          messengerBadge.style.display = 'none';
          messengerBadge.textContent = '0';
        }
      }
      // Update General Chat badge inside contacts list
      const generalBadge = document.querySelector('.contact-item[data-chat="all"] .contact-badge');
      if (generalBadge) {
        const gen = unreadCounts['all'] || 0;
        if (gen > 0) {
          generalBadge.textContent = String(gen);
          generalBadge.style.display = '';
        } else {
          generalBadge.style.display = 'none';
          generalBadge.textContent = '0';
        }
      }
      MessengerFeatureInstance.setUnreadCounts(unreadCounts);
    }
    async function ensureAdminUid() {
      if (ADMIN_UID) return ADMIN_UID;
      try {
        const snap = await userService.findByEmail(ADMIN_EMAIL);
        if (!snap.empty) {
          ADMIN_UID = snap.docs[0].id;
        }
      } catch (err) { console.error('Failed to fetch admin uid', err); }
      return ADMIN_UID;
    }

    // Old chat button functions removed - using Facebook Messenger functions instead

    // populate recipients now handled by updateContactsList(); legacy refreshRecipients removed

    // Render Facebook-style message bubble
    function appendMsg(m, canDelete = false) {
      const self = auth.currentUser?.uid === m.fromId;
      const messageGroup = document.createElement('div');
      messageGroup.className = 'message-group';

      // Format timestamp
      let ts = '';
      if (m.timestamp) {
        const d = m.timestamp.toDate ? m.timestamp.toDate() : (m.timestamp.seconds ? new Date(m.timestamp.seconds * 1000) : new Date(m.timestamp));
        ts = d.toLocaleString(undefined, { hour: '2-digit', minute: '2-digit', hour12: false });
      }

      // Show sender name for received messages in group chats
      if (!self && currentChat === 'all') {
        const senderDiv = document.createElement('div');
        senderDiv.className = 'message-sender';
        senderDiv.textContent = m.fromEmail;
        messageGroup.appendChild(senderDiv);
      }

      const bubble = document.createElement('div');
      bubble.className = `message-bubble ${self ? 'sent' : 'received'}`;
      bubble.textContent = m.text;
      bubble.dataset.id = m.id || '';

      if (canDelete && self) {
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'btn btn-sm btn-link text-danger position-absolute';
        deleteBtn.style.top = '0';
        deleteBtn.style.right = '-25px';
        deleteBtn.innerHTML = '<i class="fa fa-trash" style="font-size:10px;"></i>';
        deleteBtn.title = 'Delete message';
        deleteBtn.onclick = async () => {
          if (!confirm('Delete this message?')) return;
          deleteBtn.disabled = true;
          deleteBtn.innerHTML = '<i class="fa fa-spinner fa-spin" style="font-size:10px;"></i>';

          try {
            const user = auth.currentUser;
            // Check if user owns this message
            if (m.fromId !== user.uid) {
              alert('You can only delete your own messages');
              deleteBtn.disabled = false;
              deleteBtn.innerHTML = '<i class="fa fa-trash" style="font-size:10px;"></i>';
              return;
            }

            // Try to delete the message first
            try {
              await messagingService.deleteMessage(m.id);
              // Remove the message element from UI immediately
              messageGroup.remove();
            } catch (deleteErr) {
              // If direct deletion fails due to permissions, try soft delete
              console.log('Direct delete failed, trying soft delete:', deleteErr);
              await messagingService.updateMessage(m.id, {
                deleted: true,
                deletedBy: user.uid,
                deletedAt: serverTimestamp()
              });
              // Remove the message element from UI immediately
              messageGroup.remove();
            }

          } catch (err) {
            console.error('Delete error:', err);
            alert('Unable to delete message. Please contact administrator.');
            deleteBtn.disabled = false;
            deleteBtn.innerHTML = '<i class="fa fa-trash" style="font-size:10px;"></i>';
          }
        };
        bubble.style.position = 'relative';
        bubble.appendChild(deleteBtn);
      }

      messageGroup.appendChild(bubble);

      // Add timestamp
      if (ts) {
        const timeDiv = document.createElement('div');
        timeDiv.className = 'message-time';
        timeDiv.textContent = ts;
        messageGroup.appendChild(timeDiv);
      }

      msgContainer.appendChild(messageGroup);
      if (!suppressAutoScroll) {
        msgContainer.scrollTop = msgContainer.scrollHeight;
      }
    }

    // Update contacts list with unread counts
    function updateContactsList() {
      if (!approvedUsers) return;

      contactsList.innerHTML = '';
      // Ensure General Chat is present at top
      (function addGeneralChat() {
        const genUnread = unreadCounts['all'] || 0;
        const genItem = document.createElement('div');
        genItem.className = 'contact-item';
        genItem.dataset.chat = 'all';
        genItem.innerHTML = `
        <div class="contact-avatar">
          <i class="fa fa-users"></i>
        </div>
        <div class="contact-info">
          <div class="contact-name">General Chat</div>
          <div class="contact-preview">Broadcasts</div>
        </div>
        ${genUnread > 0 ? `<div class="contact-badge">${genUnread}</div>` : ''}
      `;
        if (currentChat === 'all') { genItem.classList.add('active'); }
        genItem.addEventListener('click', () => {
          switchChat('all', 'General Chat', 'fa-users');
        });
        contactsList.appendChild(genItem);
      })();

      approvedUsers.forEach(user => {
        const contactItem = document.createElement('div');
        contactItem.className = 'contact-item';
        contactItem.dataset.chat = user.id;

        const emailKey = (user.email || '').toLowerCase();
        const unreadCount = unreadCounts[user.id] || unreadCounts[emailKey] || 0;

        contactItem.innerHTML = `
        <div class="contact-avatar">
          <i class="fa fa-user"></i>
        </div>
        <div class="contact-info">
          <div class="contact-name">${user.email}</div>
          <div class="contact-preview">Click to start private chat</div>
        </div>
        ${unreadCount > 0 ? `<div class="contact-badge">${unreadCount}</div>` : ''}
      `;

        if (user.id === currentChat) { contactItem.classList.add('active'); }
        contactItem.addEventListener('click', () => {
          switchChat(user.id, user.email, 'fa-user');
        });

        contactsList.appendChild(contactItem);
      });

      // Add admin contact if not already present
      if (![...approvedUsers].some(u => u.email.toLowerCase() === ADMIN_EMAIL_LOWER)) {
        const adminContact = document.createElement('div');
        adminContact.className = 'contact-item';
        adminContact.dataset.chat = ADMIN_EMAIL_LOWER;

        const adminUnread = Math.max(unreadCounts[ADMIN_EMAIL_LOWER] || 0, (ADMIN_UID && unreadCounts[ADMIN_UID]) || 0);

        adminContact.innerHTML = `
        <div class="contact-avatar">
          <i class="fa fa-crown"></i>
        </div>
        <div class="contact-info">
          <div class="contact-name">${ADMIN_EMAIL}</div>
          <div class="contact-preview">Administrator</div>
        </div>
        ${adminUnread > 0 ? `<div class="contact-badge">${adminUnread}</div>` : ''}
      `;

        if (ADMIN_EMAIL_LOWER === currentChat) { adminContact.classList.add('active'); }
        adminContact.addEventListener('click', () => {
          switchChat(ADMIN_EMAIL_LOWER, ADMIN_EMAIL, 'fa-crown');
        });

        contactsList.appendChild(adminContact);
      }
    }

    // --- Facebook Messenger filtering logic ---
    function renderMessages() {
      const uid = auth.currentUser.uid;
      const userEmail = auth.currentUser.email?.toLowerCase();
      msgContainer.innerHTML = '';

      allMessages.forEach(m => {
        if (currentChat === 'all') {
          if (!m.toId) { // broadcast only
            appendMsg(m, m.fromId === uid);
          }
        } else {
          // Private thread between current user and selected recipient
          const other = currentChat; // could be uid or email string
          if (m.toId === null) return; // skip broadcasts

          const fromSelf = m.fromId === uid;
          const fromOther = m.fromId === other || (m.fromEmail && m.fromEmail.toLowerCase() === other);

          // Handle both UID and email-based targeting
          const toOther = m.toId === other || (typeof m.toId === 'string' && m.toId.toLowerCase && m.toId.toLowerCase() === other);
          const toMe = m.toId === uid || (m.toId && m.toId.toLowerCase && m.toId.toLowerCase() === userEmail) || (isAdmin && m.toId === ADMIN_EMAIL_LOWER && other === ADMIN_EMAIL_LOWER);

          const a = fromSelf && toOther; // message I sent to other
          const b = fromOther && toMe; // message other sent to me

          if (a || b) {
            appendMsg(m, m.fromId === uid);
            // mark as read using sender UID; also clear email key if any
            if (m.fromId) { unreadCounts[m.fromId] = 0; }
            if (m.fromEmail) { unreadCounts[m.fromEmail.toLowerCase()] = 0; }
            updateBadge();
            updateContactsList();
          }
        }
      });
    }

    // (Removed legacy single-image openPhotoViewer; using list-aware viewer defined earlier)

    function startListening() {
      if (msgsUnsub) msgsUnsub();
      msgsCursor = null;
      const uid = auth.currentUser.uid;
      const userEmail = auth.currentUser.email?.toLowerCase();
      if (!Array.isArray(allMessages)) allMessages = [];
      let initialized = false;
      // Listen to the last N messages (descending), then render ascending
      msgsUnsub = messagingService.subscribe(snap => {
        if (!initialized) {
          allMessages = [];
          const docs = snap.docs.slice();
          msgsCursor = docs[docs.length - 1] || null; // last doc in current window
          // build list but filter to messages relevant to me
          docs.reverse().forEach(doc => {
            const m = doc.data();
            if (m.deleted) return;
            const includeForMe = (!m.toId)
              || (m.toId === uid)
              || (m.toId && m.toId.toLowerCase && m.toId.toLowerCase() === userEmail)
              || (isAdmin && m.toId === ADMIN_EMAIL_LOWER)
              || (m.fromId === uid);
            if (includeForMe) { allMessages.push({ id: doc.id, ...m }); }
          });
          updateContactsList();
          renderMessages();
          MessengerFeatureInstance.setMessages(allMessages);
          // If fewer than a full page, there are no older messages
          if (chatLoadOlderBtn) {
            if (docs.length < CHAT_PAGE_SIZE) {
              chatLoadOlderBtn.disabled = true;
              chatLoadOlderBtn.textContent = 'No more messages';
            } else {
              chatLoadOlderBtn.disabled = false;
            }
            if (chatLoadOlderSpinner) chatLoadOlderSpinner.style.display = 'none';
          }
          initialized = true;
          return;
        }
        const changes = snap.docChanges();
        let hasAnyChange = false;
        changes.forEach(change => {
          const doc = change.doc;
          const m = doc.data();
          const id = doc.id;
          const includeForMe = (!m.toId)
            || (m.toId === uid)
            || (m.toId && m.toId.toLowerCase && m.toId.toLowerCase() === userEmail)
            || (isAdmin && m.toId === ADMIN_EMAIL_LOWER)
            || (m.fromId === uid);
          const idx = allMessages.findIndex(mm => mm.id === id);
          if (change.type === 'removed' || m.deleted) {
            if (idx !== -1) { allMessages.splice(idx, 1); hasAnyChange = true; }
            return;
          }
          if (change.type === 'modified') {
            if (idx !== -1) { allMessages[idx] = { id, ...m }; hasAnyChange = true; }
            return;
          }
          if (change.type === 'added') {
            if (includeForMe) {
              allMessages.push({ id, ...m });
              hasAnyChange = true;
              if (m.fromId !== uid) {
                if (!m.toId) {
                  if (messengerSidebar.style.display === 'none' || currentChat !== 'all') {
                    unreadCounts['all'] = (unreadCounts['all'] || 0) + 1;
                    updateBadge();
                    notifyPm(m.fromEmail);
                  }
                } else {
                  const messageToMe = (m.toId === uid || (m.toId && m.toId.toLowerCase && m.toId.toLowerCase() === userEmail) || (isAdmin && m.toId === ADMIN_EMAIL_LOWER));
                  if (messageToMe) {
                    const chatKey = m.fromId;
                    if (messengerSidebar.style.display === 'none' || currentChat !== chatKey) {
                      unreadCounts[chatKey] = (unreadCounts[chatKey] || 0) + 1;
                      updateBadge();
                      updateContactsList();
                      notifyPm(m.fromEmail);
                    }
                  }
                }
                if (messengerSidebar.style.display === 'none') {
                  messengerToggle.classList.add('animate__animated', 'animate__tada');
                  setTimeout(() => messengerToggle.classList.remove('animate__animated', 'animate__tada'), 1000);
                }
              }
            }
          }
        });
        if (hasAnyChange) {
          // Keep ascending order for display
          allMessages.sort((a, b) => {
            const ta = a.timestamp?.toMillis ? a.timestamp.toMillis() : new Date(a.timestamp || 0).getTime();
            const tb = b.timestamp?.toMillis ? b.timestamp.toMillis() : new Date(b.timestamp || 0).getTime();
            return ta - tb;
          });
          updateContactsList();
          renderMessages();
          MessengerFeatureInstance.setMessages(allMessages);
        }
        // update cursor on any snapshot
        const ds = snap.docs;
        msgsCursor = ds[ds.length - 1] || msgsCursor;
      });
    }

    // Load older messages page (prepends to current list). Returns a Promise<boolean> whether any older were loaded.
    async function loadOlderMessages() {
      if (isLoadingOlder) return false;
      if (!msgsCursor) return false;
      isLoadingOlder = true;
      try {
        const uid = auth.currentUser?.uid;
        const userEmail = auth.currentUser?.email?.toLowerCase();
        const olderSnap = await messagingService.fetchOlder(msgsCursor, CHAT_PAGE_SIZE);
        if (olderSnap.empty) { return false; }
        const docs = olderSnap.docs.slice();
        msgsCursor = docs[docs.length - 1] || msgsCursor;
        docs.reverse().forEach(doc => {
          const m = doc.data();
          if (m.deleted) return;
          const includeForMe = (!m.toId)
            || (m.toId === uid)
            || (m.toId && m.toId.toLowerCase && m.toId.toLowerCase() === userEmail)
            || (isAdmin && m.toId === ADMIN_EMAIL_LOWER)
            || (m.fromId === uid);
          if (includeForMe) {
            if (!allMessages.some(mm => mm.id === doc.id)) {
              allMessages.unshift({ id: doc.id, ...m });
            }
          }
        });
        allMessages.sort((a, b) => {
          const ta = a.timestamp?.toMillis ? a.timestamp.toMillis() : new Date(a.timestamp || 0).getTime();
          const tb = b.timestamp?.toMillis ? b.timestamp.toMillis() : new Date(b.timestamp || 0).getTime();
          return ta - tb;
        });
        renderMessages();
        return true;
      } catch (err) {
        console.error('loadOlderMessages error', err);
        return false;
      } finally {
        isLoadingOlder = false;
      }
    }
    window.loadOlderMessages = loadOlderMessages;

    // Send message function for Facebook Messenger
    async function sendMessage() {
      const txt = msgInput.value.trim();
      if (!txt) return;
      const user = auth.currentUser;

      // optimistic append
      appendMsg({ text: txt, fromId: user.uid, fromEmail: user.email, toId: currentChat === 'all' ? null : currentChat, timestamp: new Date() }, true);
      msgInput.value = '';
      try { autoResize(); } catch (_) { }

      try {
        await messagingService.sendMessage({
          text: txt,
          fromId: user.uid,
          fromEmail: user.email,
          toId: currentChat === 'all' ? null : currentChat,
          timestamp: serverTimestamp()
        });
      } catch (err) {
        alert('Failed to send: ' + err.message);
        console.error('sendMessage error', err);
      }
    }

    // Event handlers for Facebook Messenger
    messengerToggle.onclick = () => {
      showMessenger();
      messengerToggle.classList.remove('animate__animated', 'animate__tada');
    };

    messengerClose.onclick = hideMessenger;

    msgSend.onclick = sendMessage;

    msgInput.addEventListener('keyup', e => {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMessage(); }
    });
    // Auto-resize textarea height
    function autoResize() {
      if (!msgInput) return;
      msgInput.style.height = 'auto';
      const max = 120; // px max height
      msgInput.style.height = Math.min(max, msgInput.scrollHeight) + 'px';
    }
    msgInput.addEventListener('input', autoResize);
    // Initialize height
    autoResize();

    // Load older button handler with spinner and scroll position preservation
    function setOlderLoading(on) {
      if (chatLoadOlderBtn) chatLoadOlderBtn.disabled = true;
      if (chatLoadOlderSpinner) chatLoadOlderSpinner.style.display = on ? '' : 'none';
    }
    function markNoMore() {
      if (chatLoadOlderBtn) {
        chatLoadOlderBtn.disabled = true;
        chatLoadOlderBtn.textContent = 'No more messages';
      }
      if (chatLoadOlderSpinner) chatLoadOlderSpinner.style.display = 'none';
    }
    if (chatLoadOlderBtn) {
      chatLoadOlderBtn.addEventListener('click', async () => {
        if (isLoadingOlder) { return; }
        if (!msgsCursor) { markNoMore(); return; }
        const prevHeight = msgContainer.scrollHeight;
        setOlderLoading(true);
        suppressAutoScroll = true;
        const had = await loadOlderMessages();
        suppressAutoScroll = false;
        setOlderLoading(false);
        if (had) {
          const newHeight = msgContainer.scrollHeight;
          // keep viewport anchored at the same message after prepending
          msgContainer.scrollTop = newHeight - prevHeight;
        } else {
          markNoMore();
        }
      });
    }

    // General chat contact click handler
    const generalContact = document.querySelector('.contact-item[data-chat="all"]');
    if (generalContact) {
      generalContact.addEventListener('click', () => {
        switchChat('all', 'General Chat', 'fa-users');
      });
    }

    // expose for other functions
    window.__refreshChatRecipients = updateContactsList;

    // Show messenger when user logged in & approved
    function showMessengerBtn() { messengerToggle.style.display = 'inline-flex'; }
    function hideMessengerBtn() { messengerToggle.style.display = 'none'; }

    auth.onAuthStateChanged(user => {
      if (user && !user.isAnonymous) {
        showMessengerBtn();
        startListening();
        // Resolve admin UID for unread mapping and refresh contacts
        ensureAdminUid().then(updateContactsList).catch(() => { });
      } else {
        hideMessengerBtn();
        if (msgsUnsub) msgsUnsub();
      }
    });

    // Admin clear messages functionality
    if (messengerClear) {
      messengerClear.addEventListener('click', async () => {
        if (!isAdmin) return;
        if (!confirm('Delete ALL chat messages?')) return;
        messengerClear.disabled = true;
        try {
          const snap = await messagingService.fetchLatest(500);
          const batch = db.batch();
          snap.docs.forEach(doc => batch.delete(doc.ref));
          await batch.commit();
          msgContainer.innerHTML = '';
        } catch (err) {
          alert('Failed to clear messages: ' + err.message);
          console.error('Clear chat error', err);
        } finally {
          messengerClear.disabled = false;
        }
      });
    }
  })();

  // Page & modal scroll buttons
  try { window.UIShell?.initScrollControls(); } catch (_) { /* noop */ }
  // View-only link
  if (viewOnlyLink) {
    viewOnlyLink.addEventListener('click', async e => {
      e.preventDefault();
      isViewOnly = true;
      isAdmin = false;
      isLevel2 = false;
      elevatedAccess = false;
      try {
        if (auth.currentUser && !auth.currentUser.isAnonymous) {
          await auth.signOut();
        }
        await auth.signInAnonymously();
      } catch (err) {
        console.error('Anon sign-in failed', err);
        console.warn('Anon sign-in unavailable, proceeding without auth');
        isViewOnly = true;
        isAdmin = false;
        isLevel2 = false;
        elevatedAccess = false;
        showApp();
        updateAdminUI();
        updateViewOnlyUI();
        try { window.__setViewerProfile?.(); } catch (_) { /* noop */ }
        subscribeDeepwells();
        // Also subscribe to public read modules in viewer fallback
        subscribeServiceUpdates();
        subscribePresentations();
        if (unsubscribeProjects) unsubscribeProjects();
        unsubscribeProjects = projectService.subscribe(snap => {
          ProjectsFeatureInstance.setProjects(snap.docs.map(d => ({ id: d.id, ...d.data() })));
          renderProjects();
          try { renderDashboard(); } catch (_) { /* noop */ }
        });
      }
    });
  }

  // Attach listeners for other modules handled within their features
  // ---- Utility ----
  // Compress image using canvas to keep Firestore doc size small (<1MB total)
  async function compressImage(file, maxSize = 1024, quality = 0.7) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      const reader = new FileReader();
      reader.onload = e => {
        img.onload = () => {
          let { width, height } = img;
          if (width > maxSize || height > maxSize) {
            const ratio = Math.min(maxSize / width, maxSize / height);
            width = Math.round(width * ratio);
            height = Math.round(height * ratio);
          }
          const canvas = document.createElement('canvas');
          canvas.width = width; canvas.height = height;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, width, height);
          canvas.toBlob(blob => {
            const fr = new FileReader();
            fr.onload = () => resolve(fr.result);
            fr.onerror = reject;
            fr.readAsDataURL(blob);
          }, 'image/jpeg', quality);
        };
        img.onerror = reject;
        img.src = e.target.result;
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  // ---- Product Presentations (Module 6) ----
  function subscribePresentations() {
    if (unsubscribePresentations) unsubscribePresentations();
    if (!presentationService) return;
    unsubscribePresentations = presentationService.subscribe(snap => {
      presentations = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      PresentationsFeatureInstance.setPresentations(presentations);
      try { renderPresentations(); } catch (_) { }
      try { renderDashboard(); } catch (_) { }
    }, { includeMetadataChanges: true }, err => console.error('Presentations snapshot error', err));
  }

  function renderPresentations() {
    PresentationsFeatureInstance.renderPresentations();
  }

  function exportPresentationsCsv() {
    const list = PresentationsFeatureInstance.getPresentations();
    if (list.length === 0) { alert('No data to export.'); return; }
    const headers = ['Date', 'Subject', 'Time', 'Venue', 'Presenter', 'Remarks'];
    const rows = list.map(p => [
      p.date || '',
      p.subject || '',
      p.time || '',
      p.venue || '',
      p.presenter || '',
      p.remarks || ''
    ]);
    const csv = [headers.join(','), ...rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(','))].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const filename = `product_presentations_${todayYmd()}.csv`;
    if (window.saveAs) {
      window.saveAs(blob, filename);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  }

  // Presentations event listeners
  if (elements.presentationStartDateFilter) elements.presentationStartDateFilter.addEventListener('change', () => renderPresentations());
  if (elements.presentationEndDateFilter) elements.presentationEndDateFilter.addEventListener('change', () => renderPresentations());
  if (elements.presentationYearFilter) {
    elements.presentationYearFilter.addEventListener('change', () => {
      const y = elements.presentationYearFilter.value;
      if (y && y !== 'ALL') {
        if (elements.presentationStartDateFilter) elements.presentationStartDateFilter.value = jan1OfYear(parseInt(y, 10));
        if (elements.presentationEndDateFilter) {
          const isCurrent = parseInt(y, 10) === new Date().getFullYear();
          elements.presentationEndDateFilter.value = isCurrent ? todayYmd() : dec31OfYear(parseInt(y, 10));
        }
      } else if (y === 'ALL') {
        if (elements.presentationStartDateFilter) elements.presentationStartDateFilter.value = '';
        if (elements.presentationEndDateFilter) elements.presentationEndDateFilter.value = '';
      }
      renderPresentations();
    });
  }

  // Text search with debounce
  if (elements.presentationSubjectFilter) {
    const onType = debounce(() => renderPresentations(), 200);
    elements.presentationSubjectFilter.addEventListener('input', onType);
    elements.presentationSubjectFilter.addEventListener('change', () => renderPresentations());
  }
  if (elements.presentationPresenterFilter) {
    elements.presentationPresenterFilter.addEventListener('change', () => renderPresentations());
    const onType = debounce(() => renderPresentations(), 200);
    elements.presentationPresenterFilter.addEventListener('input', onType);
  }



  // Format number to 2 decimals consistently (global helper)
  function fmt2(n) {
    if (n === null || n === undefined) return '0.00';
    const num = (typeof n === 'number') ? n : parseFloat(n.toString().replace(/,/g, ''));
    if (!Number.isFinite(num)) return '0.00';
    return (Math.round(num * 100) / 100).toFixed(2);
  }

  // Format number with thousand separators and 2 decimals (e.g., 8324 -> 8,324.00)
  function fmt2c(n) {
    if (n === null || n === undefined) return '0.00';
    const num = (typeof n === 'number') ? n : parseFloat(n.toString().replace(/,/g, ''));
    if (!Number.isFinite(num)) return '0.00';
    return num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // ---- Deepwell CRUD & Rendering ----
  function bulletizeActivities(text) {
    if (!text) return '';
    // Split by patterns like "1." "2." etc. or semicolons/double line breaks
    const parts = text.split(/\s*\d+\.\s*|\s*;\s*|\n+/).filter(p => p.trim());
    if (parts.length <= 1) {
      // If splitting didn't make sense, just return original
      return text;
    }
    return `<ul class="mb-0 ps-3">${parts.map(p => `<li>${p.trim()}</li>`).join('')}</ul>`;
  }

  // ---- Deepwell CRUD & Rendering ----
  // ---- Deepwells (delegated to DeepwellsFeature) ----
  function renderDeepwells() { DeepwellsFeatureInstance.renderDeepwells(); }
  function renderDeepwellMonthlyChart() { DeepwellsFeatureInstance.renderDeepwellMonthlyChart(); }

  // Tab switching
  function showProjectsSection() {
    const mainContent = document.getElementById('mainContent');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      if (elements.dashboardTab) elements.dashboardTab.classList.remove('active');
      elements.projectsTab.classList.add('active');
      elements.deepwellsTab.classList.remove('active');
      if (elements.reforestationTab) elements.reforestationTab.classList.remove('active');
      if (elements.serviceUpdateTab) elements.serviceUpdateTab.classList.remove('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.remove('active');
      if (elements.calendarTab) elements.calendarTab.classList.remove('active');
      if (elements.opcrTab) elements.opcrTab.classList.remove('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.remove('active');

      elements.projectsSection.style.display = 'block';
      elements.projectsSection.classList.add('fade-element');
      elements.deepwellsSection.style.display = 'none';
      if (elements.reforestationSection) elements.reforestationSection.style.display = 'none';
      if (elements.serviceUpdateSection) elements.serviceUpdateSection.style.display = 'none';
      if (elements.presentationsSection) elements.presentationsSection.style.display = 'none';
      if (elements.calendarSection) elements.calendarSection.style.display = 'none';
      if (elements.opcrSection) elements.opcrSection.style.display = 'none';
      if (elements.opcrFormSection) elements.opcrFormSection.style.display = 'none';
      if (elements.ipcrSection) elements.ipcrSection.style.display = 'none';
      if (elements.dashboardSection) elements.dashboardSection.style.display = 'none';
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';

      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }
  function showDeepwellsSection() {
    const mainContent = document.getElementById('mainContent');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      if (elements.dashboardTab) elements.dashboardTab.classList.remove('active');
      elements.projectsTab.classList.remove('active');
      elements.deepwellsTab.classList.add('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.remove('active');
      if (elements.calendarTab) elements.calendarTab.classList.remove('active');
      if (elements.opcrTab) elements.opcrTab.classList.remove('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.remove('active');

      elements.projectsSection.style.display = 'none';
      elements.deepwellsSection.style.display = 'block';
      elements.deepwellsSection.classList.add('fade-element');
      if (elements.reforestationSection) elements.reforestationSection.style.display = 'none';
      if (elements.serviceUpdateSection) elements.serviceUpdateSection.style.display = 'none';
      if (elements.presentationsSection) elements.presentationsSection.style.display = 'none';
      if (elements.calendarSection) elements.calendarSection.style.display = 'none';
      if (elements.opcrSection) elements.opcrSection.style.display = 'none';
      if (elements.opcrFormSection) elements.opcrFormSection.style.display = 'none';
      if (elements.ipcrSection) elements.ipcrSection.style.display = 'none';
      if (elements.dashboardSection) elements.dashboardSection.style.display = 'none';
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';

      renderDeepwells();
      renderDeepwellMonthlyChart();

      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }

  // Show Dashboard section
  function showDashboardSection() {
    const mainContent = document.getElementById('mainContent');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      // set active tab state
      if (elements.dashboardTab) elements.dashboardTab.classList.add('active');
      elements.projectsTab.classList.remove('active');
      elements.deepwellsTab.classList.remove('active');
      if (elements.reforestationTab) elements.reforestationTab.classList.remove('active');
      if (elements.serviceUpdateTab) elements.serviceUpdateTab.classList.remove('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.remove('active');
      if (elements.calendarTab) elements.calendarTab.classList.remove('active');
      if (elements.opcrTab) elements.opcrTab.classList.remove('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.remove('active');

      // toggle sections
      if (elements.dashboardSection) {
        elements.dashboardSection.style.display = 'block';
        elements.dashboardSection.classList.add('fade-element');
      }
      elements.projectsSection.style.display = 'none';
      elements.deepwellsSection.style.display = 'none';
      if (elements.reforestationSection) elements.reforestationSection.style.display = 'none';
      if (elements.serviceUpdateSection) elements.serviceUpdateSection.style.display = 'none';
      if (elements.presentationsSection) elements.presentationsSection.style.display = 'none';
      if (elements.calendarSection) elements.calendarSection.style.display = 'none';
      if (elements.opcrSection) elements.opcrSection.style.display = 'none';
      if (elements.opcrFormSection) elements.opcrFormSection.style.display = 'none';
      if (elements.ipcrSection) elements.ipcrSection.style.display = 'none';
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';

      // render dashboard widgets/charts
      try { renderDashboard(); } catch (_) { /* noop */ }


      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }

  elements.projectsTab?.addEventListener('click', e => { e.preventDefault(); showProjectsSection(); });
  elements.deepwellsTab?.addEventListener('click', e => { e.preventDefault(); showDeepwellsSection(); });
  elements.dashboardTab?.addEventListener('click', e => { e.preventDefault(); showDashboardSection(); });


  function showServiceUpdateSection() {
    const mainContent = document.getElementById('mainContent');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      if (elements.dashboardTab) elements.dashboardTab.classList.remove('active');
      elements.projectsTab.classList.remove('active');
      elements.deepwellsTab.classList.remove('active');
      if (elements.reforestationTab) elements.reforestationTab.classList.remove('active');
      if (elements.serviceUpdateTab) elements.serviceUpdateTab.classList.add('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.remove('active');
      if (elements.calendarTab) elements.calendarTab.classList.remove('active');
      if (elements.opcrTab) elements.opcrTab.classList.remove('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.remove('active');

      elements.projectsSection.style.display = 'none';
      elements.deepwellsSection.style.display = 'none';
      if (elements.reforestationSection) elements.reforestationSection.style.display = 'none';
      if (elements.serviceUpdateSection) {
        elements.serviceUpdateSection.style.display = 'block';
        elements.serviceUpdateSection.classList.add('fade-element');
      }
      if (elements.presentationsSection) elements.presentationsSection.style.display = 'none';
      if (elements.calendarSection) elements.calendarSection.style.display = 'none';
      if (elements.opcrSection) elements.opcrSection.style.display = 'none';
      if (elements.opcrFormSection) elements.opcrFormSection.style.display = 'none';
      if (elements.ipcrSection) elements.ipcrSection.style.display = 'none';
      if (elements.dashboardSection) elements.dashboardSection.style.display = 'none';
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';

      if (typeof renderServiceUpdates === 'function') try { renderServiceUpdates(); } catch (_) {/*noop*/ }
      if (typeof renderSurChart === 'function') try { renderSurChart(); } catch (_) {/*noop*/ }

      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }
  elements.serviceUpdateTab?.addEventListener('click', e => { e.preventDefault(); showServiceUpdateSection(); });

  // Presentations section switching
  function showPresentationsSection() {
    const mainContent = document.getElementById('mainContent');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      if (elements.dashboardTab) elements.dashboardTab.classList.remove('active');
      elements.projectsTab.classList.remove('active');
      elements.deepwellsTab.classList.remove('active');
      if (elements.reforestationTab) elements.reforestationTab.classList.remove('active');
      if (elements.serviceUpdateTab) elements.serviceUpdateTab.classList.remove('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.add('active');
      if (elements.calendarTab) elements.calendarTab.classList.remove('active');
      if (elements.opcrTab) elements.opcrTab.classList.remove('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.remove('active');

      elements.projectsSection.style.display = 'none';
      elements.deepwellsSection.style.display = 'none';
      if (elements.reforestationSection) elements.reforestationSection.style.display = 'none';
      if (elements.serviceUpdateSection) elements.serviceUpdateSection.style.display = 'none';
      if (elements.dashboardSection) elements.dashboardSection.style.display = 'none';
      if (elements.presentationsSection) {
        elements.presentationsSection.style.display = 'block';
        elements.presentationsSection.classList.add('fade-element');
      }
      if (elements.calendarSection) elements.calendarSection.style.display = 'none';
      if (elements.opcrSection) elements.opcrSection.style.display = 'none';
      if (elements.opcrFormSection) elements.opcrFormSection.style.display = 'none';
      if (elements.ipcrSection) elements.ipcrSection.style.display = 'none';
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';

      try { renderPresentations(); } catch (_) {/* noop */ }

      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }
  elements.presentationsTab?.addEventListener('click', e => { e.preventDefault(); showPresentationsSection(); });

  // OPCR section switching (Apps)
  function showOpcrSection() {
    const mainContent = document.getElementById('mainContent');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      if (elements.dashboardTab) elements.dashboardTab.classList.remove('active');
      elements.projectsTab.classList.remove('active');
      elements.deepwellsTab.classList.remove('active');
      if (elements.reforestationTab) elements.reforestationTab.classList.remove('active');
      if (elements.serviceUpdateTab) elements.serviceUpdateTab.classList.remove('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.remove('active');
      if (elements.calendarTab) elements.calendarTab.classList.remove('active');
      if (elements.opcrTab) elements.opcrTab.classList.add('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.remove('active');

      // Hide all sections
      elements.projectsSection.style.display = 'none';
      elements.deepwellsSection.style.display = 'none';
      if (elements.reforestationSection) elements.reforestationSection.style.display = 'none';
      if (elements.serviceUpdateSection) elements.serviceUpdateSection.style.display = 'none';
      if (elements.dashboardSection) elements.dashboardSection.style.display = 'none';
      if (elements.presentationsSection) elements.presentationsSection.style.display = 'none';
      if (elements.calendarSection) elements.calendarSection.style.display = 'none';
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';
      if (elements.opcrFormSection) elements.opcrFormSection.style.display = 'none';
      if (elements.ipcrSection) elements.ipcrSection.style.display = 'none';

      // Show OPCR
      if (elements.opcrSection) {
        elements.opcrSection.style.display = 'block';
        elements.opcrSection.classList.add('fade-element');
      }

      try { renderOpcr(); } catch (_) { /* noop */ }

      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }
  elements.headerOpcrBtn?.addEventListener('click', e => {
    e.preventDefault();
    const appsMenu = document.getElementById('appsMenu');
    if (appsMenu) appsMenu.style.display = 'none';
    showOpcrSection();
  });

  function renderOpcr() { OpcrFeatureInstance.renderOpcr(); }
  function subscribeOpcr() {
    if (unsubscribeOpcr) unsubscribeOpcr();
    unsubscribeOpcr = opcrService.subscribe(snap => {
      opcrEntries = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      OpcrFeatureInstance.setOpcrEntries(opcrEntries);
      try { renderOpcr(); } catch (_) {/*noop*/ }
    });
  }

  // IPCR section switching
  let ipcrEntries = [];
  let unsubscribeIpcr = null;

  function showIpcrSection() {
    const mainContent = document.getElementById('mainContent');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      if (elements.dashboardTab) elements.dashboardTab.classList.remove('active');
      elements.projectsTab.classList.remove('active');
      elements.deepwellsTab.classList.remove('active');
      if (elements.reforestationTab) elements.reforestationTab.classList.remove('active');
      if (elements.serviceUpdateTab) elements.serviceUpdateTab.classList.remove('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.remove('active');
      if (elements.calendarTab) elements.calendarTab.classList.remove('active');
      if (elements.opcrTab) elements.opcrTab.classList.remove('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.add('active');

      // Hide all sections
      elements.projectsSection.style.display = 'none';
      elements.deepwellsSection.style.display = 'none';
      if (elements.reforestationSection) elements.reforestationSection.style.display = 'none';
      if (elements.serviceUpdateSection) elements.serviceUpdateSection.style.display = 'none';
      if (elements.dashboardSection) elements.dashboardSection.style.display = 'none';
      if (elements.presentationsSection) elements.presentationsSection.style.display = 'none';
      if (elements.calendarSection) elements.calendarSection.style.display = 'none';
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';
      if (elements.opcrSection) elements.opcrSection.style.display = 'none';
      if (elements.opcrFormSection) elements.opcrFormSection.style.display = 'none';

      // Show IPCR
      if (elements.ipcrSection) {
        elements.ipcrSection.style.display = 'block';
        elements.ipcrSection.classList.add('fade-element');
      }

      try { renderIpcr(); } catch (_) { /* noop */ }

      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }
  elements.headerIpcrBtn?.addEventListener('click', e => {
    e.preventDefault();
    const appsMenu = document.getElementById('appsMenu');
    if (appsMenu) appsMenu.style.display = 'none';
    showIpcrSection();
  });

  function renderIpcr() { IpcrFeatureInstance.renderIpcr(); }
  function subscribeIpcr() {
    if (unsubscribeIpcr) unsubscribeIpcr();
    unsubscribeIpcr = ipcrService.subscribe(snap => {
      ipcrEntries = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      IpcrFeatureInstance.setIpcrEntries(ipcrEntries);
      try { renderIpcr(); } catch (_) {/*noop*/ }
    });
  }


  // Calendar Section
  function showCalendarSection() {
    const mainContent = document.getElementById('mainContent');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      if (elements.dashboardTab) elements.dashboardTab.classList.remove('active');
      elements.projectsTab.classList.remove('active');
      elements.deepwellsTab.classList.remove('active');
      if (elements.reforestationTab) elements.reforestationTab.classList.remove('active');
      if (elements.serviceUpdateTab) elements.serviceUpdateTab.classList.remove('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.remove('active');
      if (elements.calendarTab) elements.calendarTab.classList.add('active');
      if (elements.opcrTab) elements.opcrTab.classList.remove('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.remove('active');

      // Hide all sections
      if (elements.dashboardSection) elements.dashboardSection.style.display = 'none';
      elements.projectsSection.style.display = 'none';
      elements.deepwellsSection.style.display = 'none';
      if (elements.reforestationSection) elements.reforestationSection.style.display = 'none';
      if (elements.serviceUpdateSection) elements.serviceUpdateSection.style.display = 'none';
      if (elements.presentationsSection) elements.presentationsSection.style.display = 'none';
      if (elements.opcrSection) elements.opcrSection.style.display = 'none';
      if (elements.opcrFormSection) elements.opcrFormSection.style.display = 'none';
      if (elements.ipcrSection) elements.ipcrSection.style.display = 'none';
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';

      // Show Calendar
      if (elements.calendarSection) {
        elements.calendarSection.style.display = 'block';
        elements.calendarSection.classList.add('fade-element');
      }

      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }
  elements.calendarTab?.addEventListener('click', e => { e.preventDefault(); showCalendarSection(); });

  // Document Registry: show section in-page
  function showDocumentRegistrySection() {
    const mainContent = document.getElementById('mainContent');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      // Remove active from all tabs
      if (elements.dashboardTab) elements.dashboardTab.classList.remove('active');
      elements.projectsTab.classList.remove('active');
      elements.deepwellsTab.classList.remove('active');
      if (elements.reforestationTab) elements.reforestationTab.classList.remove('active');
      if (elements.serviceUpdateTab) elements.serviceUpdateTab.classList.remove('active');
      if (elements.presentationsTab) elements.presentationsTab.classList.remove('active');
      if (elements.calendarTab) elements.calendarTab.classList.remove('active');
      if (elements.opcrTab) elements.opcrTab.classList.remove('active');
      if (elements.ipcrTab) elements.ipcrTab.classList.remove('active');

      // Hide all sections
      elements.projectsSection.style.display = 'none';
      elements.deepwellsSection.style.display = 'none';
      if (elements.reforestationSection) elements.reforestationSection.style.display = 'none';
      if (elements.serviceUpdateSection) elements.serviceUpdateSection.style.display = 'none';
      if (elements.dashboardSection) elements.dashboardSection.style.display = 'none';
      if (elements.presentationsSection) elements.presentationsSection.style.display = 'none';
      if (elements.calendarSection) elements.calendarSection.style.display = 'none';
      if (elements.personalTasksSection) elements.personalTasksSection.style.display = 'none';
      if (elements.documentRegistrySection) elements.documentRegistrySection.style.display = 'none';
      if (elements.opcrSection) elements.opcrSection.style.display = 'none';
      if (elements.opcrFormSection) elements.opcrFormSection.style.display = 'none';
      if (elements.ipcrSection) elements.ipcrSection.style.display = 'none';

      // Show Document Registry section
      if (elements.documentRegistrySection) {
        elements.documentRegistrySection.style.display = 'block';
        elements.documentRegistrySection.classList.add('fade-element');
      }

      // Update global header title
      const headerTitle = document.getElementById('globalModuleTitle');
      if (headerTitle) headerTitle.textContent = 'Document Registry';

      // Initialize if docreg.js provides an init function
      if (typeof window.initDocReg === 'function') {
        window.initDocReg();
      }

      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }
  if (elements.documentRegistryBtn) {
    elements.documentRegistryBtn.addEventListener('click', () => { showDocumentRegistrySection(); });
  }
  // Personal Tasks: UI wiring
  if (elements.personalTasksBtn) {
    elements.personalTasksBtn.addEventListener('click', showPersonalTasksSection);
  }
  if (elements.ptmAddBtn) {
    elements.ptmAddBtn.addEventListener('click', () => { clearPersonalTaskForm(); elements.personalTaskModal.show(); });
  }
  if (elements.personalTaskForm) { elements.personalTaskForm.addEventListener('submit', savePersonalTask); }
  if (elements.ptmAddSubtask) { elements.ptmAddSubtask.addEventListener('click', () => addSubtaskRow()); }
  if (elements.ptmProgress && elements.ptmProgressLabel) { elements.ptmProgress.addEventListener('input', () => { elements.ptmProgressLabel.textContent = `${elements.ptmProgress.value}%`; }); }
  // Auto-set progress to 100% when status is changed to Done
  if (elements.personalTaskForm) {
    const statusSel = elements.personalTaskForm.querySelector('#ptmStatus');
    if (statusSel && elements.ptmProgress && elements.ptmProgressLabel) {
      statusSel.addEventListener('change', () => {
        if (statusSel.value === 'Done') {
          elements.ptmProgress.value = 100;
          elements.ptmProgressLabel.textContent = '100%';
          // If Date Completed empty, set to today by default for convenience
          const dcEl = elements.personalTaskForm.querySelector('#ptmDateCompleted');
          if (dcEl) {
            if (!dcEl.value) { dcEl.value = new Date().toISOString().slice(0, 10); }
            dcEl.required = true;
          }
        } else {
          // Not Done: clear Date Completed and remove required
          const dcEl = elements.personalTaskForm.querySelector('#ptmDateCompleted');
          if (dcEl) { dcEl.value = ''; dcEl.required = false; }
        }
      });
    }
  }
  if (elements.ptmCategoryFilter) { elements.ptmCategoryFilter.addEventListener('change', () => { personalTasksPage = 1; renderPersonalTasks(); }); }
  if (elements.ptmStatusFilter) { elements.ptmStatusFilter.addEventListener('change', () => { personalTasksPage = 1; renderPersonalTasks(); }); }
  if (elements.ptmPriorityFilter) { elements.ptmPriorityFilter.addEventListener('change', () => { personalTasksPage = 1; renderPersonalTasks(); }); }
  if (elements.ptmTitleSearch) { elements.ptmTitleSearch.addEventListener('input', () => { personalTasksPage = 1; renderPersonalTasks(); }); }


  // ---- Service Updates (delegated to ServiceUpdatesFeature) ----
  function renderServiceUpdates() { ServiceUpdatesFeatureInstance.renderServiceUpdates(); }

  // ---- Service Update Report CRUD & Rendering ----
  // Plant lists by concessionaire (exclude Supply Augment which is tracked separately)
  function getProviderPlants(provider) {
    return (provider || 'MWCI') === 'MWCI' ? MWCI_PLANTS : MWSI_PLANTS;
  }

  function renderSurPlantTable(provider, plantsData) {
    const container = document.getElementById('surPlantTableContainer');
    if (!container) return;
    const list = getProviderPlants(provider);
    // build rows
    let html = '';
    html += '<div class="table-responsive">';
    html += '<table class="table table-sm align-middle mb-2">';
    html += '<thead><tr class="table-light"><th>Plant</th><th style="width:160px">Raw Inflows (MLD)</th><th style="width:200px">Daily Production (MLD)</th></tr></thead><tbody>';
    for (let i = 0; i < list.length; i++) {
      const name = list[i];
      const existing = Array.isArray(plantsData) ? (plantsData.find(p => p.name === name) || {}) : {};
      const inflow = typeof existing.inflow === 'number' ? existing.inflow : 0;
      const prod = typeof existing.production === 'number' ? existing.production : 0;
      html += `<tr>`
        + `<td class="small">${name}</td>`
        + `<td><input type="number" step="0.01" class="form-control form-control-sm" data-role="inflow" data-plant="${name}" value="${inflow}"></td>`
        + `<td><input type="number" step="0.01" class="form-control form-control-sm" data-role="production" data-plant="${name}" value="${prod}"></td>`
        + `</tr>`;
    }
    html += '</tbody></table>';
    html += '</div>';
    container.innerHTML = html;
    // attach input listener for auto-sum (guard to prevent duplicates across re-renders)
    if (!container.__wired) { container.addEventListener('input', updateSurTotals); container.__wired = true; }
    // compute initial totals
    updateSurTotals();
  }

  function toYmdLocal(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function prevDateStr(ymdStr) {
    if (!ymdStr) return '';
    const parts = ymdStr.split('-').map(n => parseInt(n, 10));
    if (parts.length !== 3 || parts.some(isNaN)) return '';
    const dt = new Date(parts[0], parts[1] - 1, parts[2]);
    dt.setDate(dt.getDate() - 1);
    return toYmdLocal(dt);
  }

  function updateSurPlantsDateLabel() {
    const label = document.getElementById('surPlantsDateLabel');
    if (!label) return;
    const dateEl = document.getElementById('surDate');
    const v = dateEl ? dateEl.value : '';
    label.textContent = v || '';
  }

  function updateSurDamDateLabel() {
    const label = document.getElementById('surDamDateLabel');
    if (!label) return;
    const dateEl = document.getElementById('surDate');
    const v = dateEl ? dateEl.value : '';
    label.textContent = v || '';
  }

  function updateSurDateLabels() {
    updateSurPlantsDateLabel();
    updateSurDamDateLabel();
  }

  function readSurPlantData() {
    const container = document.getElementById('surPlantTableContainer');
    if (!container) return [];
    const rows = [];
    const inflowInputs = Array.from(container.querySelectorAll('input[data-role="inflow"]'));
    inflowInputs.forEach(input => {
      const name = input.getAttribute('data-plant');
      const inflow = parseFloat(input.value) || 0;
      const prodInput = container.querySelector(`input[data-role="production"][data-plant="${CSS.escape(name)}"]`);
      const production = parseFloat(prodInput?.value) || 0;
      rows.push({ name, inflow, production });
    });
    return rows;
  }

  function updateSurTotals() {
    const plants = readSurPlantData();
    const inflowTotal = plants.reduce((sum, p) => sum + (p.inflow || 0), 0);
    const prodTotal = plants.reduce((sum, p) => sum + (p.production || 0), 0);
    const inflowEl = document.getElementById('surInflows');
    const prodEl = document.getElementById('surProduction');
    if (inflowEl) inflowEl.value = (Math.round(inflowTotal * 100) / 100).toFixed(2);
    if (prodEl) prodEl.value = (Math.round(prodTotal * 100) / 100).toFixed(2);
  }

  // --- Excel/CSV import for Service Update per-plant metrics ---
  function normalizeHeader(h) {
    return String(h || '').trim().toLowerCase();
  }

  function parseNumberSafe(v) {
    if (typeof v === 'number') return v;
    if (typeof v === 'string') {
      const n = parseFloat(v.replace(/,/g, ''));
      return isNaN(n) ? 0 : n;
    }
    return 0;
  }

  function handleSurImportFile(evt) {
    const file = evt.target.files && evt.target.files[0];
    if (!file) { return; }
    try {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: 'array' });
          const sheetName = wb.SheetNames[0];
          const ws = wb.Sheets[sheetName];
          if (!ws) { alert('Unable to read worksheet from the uploaded file.'); return; }
          // Convert to JSON with header row auto-detected
          const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
          if (!Array.isArray(rows) || rows.length === 0) { alert('The uploaded file appears empty.'); return; }
          // Determine provider to map plants
          const f = document.getElementById('serviceUpdateForm') || document;
          const provider = (f.querySelector('#surProvider')?.value || 'MWCI');
          const validPlants = new Set(getProviderPlants(provider).map(n => n.trim().toLowerCase()))
          // Determine likely columns
          // Accept variants: Plant, Facility, Name; Inflow: Inflow, Raw Inflow(s), Raw Water Inflows; Production: Production, Daily Production, Treatment Production
          const sample = rows[0];
          const headers = Object.keys(sample || {});
          const findCol = (cands) => {
            for (const h of headers) { if (cands.includes(normalizeHeader(h))) return h; }
            return null;
          };
          const plantCol = findCol(['plant', 'facility', 'name']);
          const inflowCol = findCol(['inflow', 'raw inflow', 'raw inflows', 'raw water inflow', 'raw water inflows', 'raw water (mld)', 'raw inflows (mld)']);
          const prodCol = findCol(['production', 'daily production', 'treatment production', 'prod', 'production (mld)', 'daily production (mld)']);
          if (!plantCol) { alert('Could not detect a Plant/Facility column in the uploaded file. Please include a Plant column.'); return; }
          if (!inflowCol && !prodCol) { alert('Could not detect Inflow or Production columns. Please include at least one of them.'); return; }
          // Build plants data
          const imported = [];
          rows.forEach(r => {
            const nameRaw = (r[plantCol] || '').toString().trim();
            const key = nameRaw.toLowerCase();
            if (!nameRaw || !validPlants.has(key)) return; // skip unknown plants for selected provider
            const inflow = inflowCol ? parseNumberSafe(r[inflowCol]) : 0;
            const production = prodCol ? parseNumberSafe(r[prodCol]) : 0;
            imported.push({ name: getProviderPlants(provider).find(n => n.trim().toLowerCase() === key) || nameRaw, inflow, production });
          });
          if (imported.length === 0) { alert('No matching plant rows were found for the selected provider.'); return; }
          // Render table with imported values and auto-sum
          renderSurPlantTable(provider, imported);
          // Clear the file input so the same file can be re-selected if needed
          evt.target.value = '';
        } catch (err) {
          console.error('Import parse error', err);
          alert('Failed to parse the uploaded file. Please ensure it is a valid Excel/CSV with Plant/Inflow/Production columns.');
        }
      };
      reader.onerror = () => alert('Failed to read the uploaded file.');
      reader.readAsArrayBuffer(file);
    } catch (err) {
      alert('Import failed: ' + (err?.message || err));
    }
  }

  function downloadSurTemplate() {
    try {
      const f = document.getElementById('serviceUpdateForm') || document;
      const provider = (f.querySelector('#surProvider')?.value || 'MWCI');
      const date = f.querySelector('#surDate')?.value || fmtYmd(new Date());
      const plantList = getProviderPlants(provider);
      const rows = plantList.map(name => ({
        'Plant': name,
        'Raw Inflows (MLD)': '',
        'Daily Production (MLD)': ''
      }));
      const headers = ['Plant', 'Raw Inflows (MLD)', 'Daily Production (MLD)'];
      const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Template');
      const fname = `SUR_Template_${provider}_${date}.xlsx`;
      XLSX.writeFile(wb, fname);
    } catch (err) {
      console.error('Template generation error', err);
      alert('Failed to generate template.');
    }
  }

  // ===== Bulk Import/Export for Service Update Reports =====
  function fmtYmd(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  // Canonicalize provider values (trim, remove NBSP, uppercase, map aliases)
  function canonProvider(v) {
    const s = String(v || '').replace(/\u00A0/g, ' ').trim().toUpperCase();
    if (!s) return '';
    if (s === "MWCI" || s.includes('MANILA WATER')) return "MWCI";
    if (s === "MWSI" || s.includes('MAYNILAD')) return "MWSI";
    return s;
  }
  function prevDateStrLocal(ymd) {
    // ymd: YYYY-MM-DD
    const [y, m, d] = (ymd || '').split('-').map(x => parseInt(x, 10));
    if ([y, m, d].some(isNaN)) return '';
    const dt = new Date(y, m - 1, d);
    dt.setDate(dt.getDate() - 1);
    return fmtYmd(dt);
  }
  function eachDateInclusive(startYmd, endYmd) {
    const dates = [];
    if (!startYmd || !endYmd) return dates;
    let [sy, sm, sd] = startYmd.split('-').map(n => parseInt(n, 10));
    let [ey, em, ed] = endYmd.split('-').map(n => parseInt(n, 10));
    if ([sy, sm, sd, ey, em, ed].some(isNaN)) return dates;
    const d = new Date(sy, sm - 1, sd);
    const end = new Date(ey, em - 1, ed);
    while (d <= end) { dates.push(fmtYmd(d)); d.setDate(d.getDate() + 1); }
    return dates;
  }

  function bulkHeaderForProvider(provider) {
    const base = [
      'Date',
      'Concessionaire',
      'Dam As of',
      'Supply Augmentation (MLD)',
      'Angat Dam (masl)',
      'Ipo Dam (masl)',
      'La Mesa Dam (masl)'
    ];
    const plants = getProviderPlants(provider);
    const inflowCols = [];
    const productionCols = [];
    plants.forEach(p => { inflowCols.push(`${p} - Inflow (MLD)`); });
    plants.forEach(p => { productionCols.push(`${p} - Production (MLD)`); });
    return [...base, ...inflowCols, ...productionCols];
  }

  function downloadSurBulkTemplate() {
    try {
      const start = elements.serviceStartDateFilter?.value || '';
      const end = elements.serviceStartDateFilter?.value || '';
      const dates = eachDateInclusive(start, end);
      if (dates.length === 0) { alert('Please set a valid Start/End Date to generate the bulk template.'); return; }
      const wb = XLSX.utils.book_new();
      ['MWCI', 'MWSI'].forEach(provider => {
        const headers = bulkHeaderForProvider(provider);
        const rows = dates.map(date => {
          const r = {
            'Date': date,
            'Concessionaire': provider,
            'Dam As of': '',
            'Supply Augmentation (MLD)': '',
            'Angat Dam (masl)': '',
            'Ipo Dam (masl)': '',
            'La Mesa Dam (masl)': ''
          };
          getProviderPlants(provider).forEach(p => {
            r[`${p} - Inflow (MLD)`] = '';
            r[`${p} - Production (MLD)`] = '';
          });
          return r;
        });
        const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
        XLSX.utils.book_append_sheet(wb, ws, provider);
      });
      const fname = `SUR_Bulk_Template_${start}_to_${end}.xlsx`;
      XLSX.writeFile(wb, fname);
    } catch (err) {
      console.error('Bulk template error', err);
      alert('Failed to generate bulk template.');
    }
  }

  async function handleSurBulkImportFile(evt) {
    const file = evt.target.files && evt.target.files[0];
    if (!file) { return; }
    if (!elevatedAccess) { alert('Only admin/level2 can bulk import.'); evt.target.value = ''; return; }
    try {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: 'array' });

          // Smart sheet detection: prefer MWCI/MWSI sheets, fallback to first sheet
          let providers = wb.SheetNames.filter(n => ['mwci', 'mwsi'].includes(n.trim().toLowerCase()));

          // If no MWCI/MWSI sheets, use all sheets but try to detect provider from data
          if (providers.length === 0) {
            providers = wb.SheetNames.slice(0, 2); // Use up to first 2 sheets
          }

          if (providers.length === 0) {
            alert('No data sheets found in workbook.');
            evt.target.value = '';
            return;
          }

          let total = 0;
          const toSave = [];

          providers.forEach(sheetName => {
            const ws = wb.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
            if (!rows.length) return;

            // Detect provider from sheet name or data
            let detectedProvider = canonProvider(sheetName);
            if (!detectedProvider || (detectedProvider !== 'MWCI' && detectedProvider !== 'MWSI')) {
              // Try to detect from first row's Concessionaire column
              const firstProvider = canonProvider(rows[0]?.['Concessionaire'] || rows[0]?.['Provider'] || '');
              detectedProvider = (firstProvider === 'MWCI' || firstProvider === 'MWSI') ? firstProvider : 'MWCI';
            }

            const plantList = getProviderPlants(detectedProvider);

            rows.forEach(r => {
              const date = normalizeYmd(r['Date']);
              if (!date) return;

              // Get provider from row or use detected default
              const rowProvider = canonProvider(r['Concessionaire'] || r['Provider'] || detectedProvider);
              const prov = (rowProvider === 'MWCI' || rowProvider === 'MWSI') ? rowProvider : detectedProvider;
              const currentPlantList = getProviderPlants(prov);

              // Helper: fuzzy match column name to plant name
              const findPlantColumn = (plantName, suffix) => {
                const keys = Object.keys(r);
                // Try exact match first
                const exact = `${plantName} - ${suffix} (MLD)`;
                if (keys.includes(exact)) return r[exact];

                // Extract key words from plant name (e.g., "Balara WTP 1" -> ["balara", "wtp", "1"])
                const plantKeywords = plantName.toLowerCase().replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(w => w.length > 1);

                // Search for columns that contain the plant keywords
                for (const key of keys) {
                  const keyLower = key.toLowerCase();
                  // Check if column contains main plant identifier
                  const hasMainKeyword = plantKeywords.some(kw => keyLower.includes(kw));
                  if (hasMainKeyword) {
                    // For columns with explicit suffix, check for inflow/production
                    if (suffix && (keyLower.includes('inflow') || keyLower.includes('production'))) {
                      if (suffix.toLowerCase() === 'inflow' && keyLower.includes('inflow')) return r[key];
                      if (suffix.toLowerCase() === 'production' && keyLower.includes('production')) return r[key];
                    }
                    // For columns without suffix (direct plant values), return if it's a number
                    const val = parseNumberSafe(r[key]);
                    if (val > 0 && !keyLower.includes('inflow') && !keyLower.includes('production')) {
                      return r[key]; // Return the raw value for later parsing
                    }
                  }
                }
                return null;
              };

              const plants = [];
              currentPlantList.forEach(p => {
                // Try standard format first  
                let inflow = parseNumberSafe(r[`${p} - Inflow (MLD)`]);
                let production = parseNumberSafe(r[`${p} - Production (MLD)`]);

                // If not found, try fuzzy matching
                if (!inflow && !production) {
                  inflow = parseNumberSafe(findPlantColumn(p, 'Inflow')) || 0;
                  production = parseNumberSafe(findPlantColumn(p, 'Production')) || 0;

                  // If still not found, just try to find any column with the plant name
                  if (!inflow && !production) {
                    const directVal = findPlantColumn(p, null);
                    if (directVal !== null) {
                      // Assume direct values are production
                      production = parseNumberSafe(directVal) || 0;
                    }
                  }
                }

                if (inflow || production) { plants.push({ name: p, inflow, production }); }
              });

              // Calculate inflows/production from plants if available, otherwise use total columns (from export format)
              let inflows = plants.reduce((s, x) => s + (x.inflow || 0), 0);
              let production = plants.reduce((s, x) => s + (x.production || 0), 0);

              // If no plant data found, try reading totals directly from export format columns
              if (plants.length === 0 || (inflows === 0 && production === 0)) {
                const totalInflows = parseNumberSafe(r['Total Raw Inflows (MLD)'] || r['Raw Inflows (MLD)'] || r['Inflows']);
                const totalProduction = parseNumberSafe(r['Total Treatment Production (MLD)'] || r['Treatment Prod (MLD)'] || r['Production']);
                if (totalInflows > 0) inflows = totalInflows;
                if (totalProduction > 0) production = totalProduction;
              }

              const s = {
                date: date || '',
                provider: prov || 'MWCI',
                damAsOf: (r['Dam As of'] || r['Dam As Of'] || r['As of'] || '').toString().trim(),
                inflows: Math.round(inflows * 100) / 100 || 0,
                production: Math.round(production * 100) / 100 || 0,
                supplyAug: parseNumberSafe(r['Supply Augmentation (MLD)'] || r['Supply Aug (MLD)'] || r['Supply Aug']) || 0,
                angat: parseNumberSafe(r['Angat Dam (masl)'] || r['Angat (masl)'] || r['Angat']) || 0,
                ipo: parseNumberSafe(r['Ipo Dam (masl)'] || r['Ipo (masl)'] || r['Ipo']) || 0,
                laMesa: parseNumberSafe(r['La Mesa Dam (masl)'] || r['La Mesa (masl)'] || r['La Mesa']) || 0,
                plants: plants || [],
                plantsDate: date || '',
                remarks: (r['Remarks'] || '').toString(),
                timestamp: Date.now()
              };

              // Upsert: attach existing id if found in current cache
              const existing = serviceUpdates.find(x => (x.date || '') === date && ((x.provider || '').toUpperCase() === prov));
              if (existing) { s.id = existing.id; }
              toSave.push(s);
            });
          });

          total = toSave.length;
          if (total === 0) {
            alert('No valid rows found to import. Please ensure your file has a Date column and proper format.');
            evt.target.value = '';
            return;
          }

          if (!confirm(`Import ${total} report(s)? This will create/update records by date & provider.`)) {
            evt.target.value = '';
            return;
          }

          // Save in parallel with detailed error handling
          const results = await Promise.all(toSave.map(async (item, idx) => {
            try {
              await saveServiceUpdate(item);
              return { success: true };
            } catch (saveErr) {
              console.error(`Failed to save record ${idx + 1}:`, item, saveErr);
              return { success: false, index: idx, date: item?.date, error: saveErr?.message || String(saveErr) };
            }
          }));

          const failedRecords = results.filter(r => !r.success);
          const savedCount = results.length - failedRecords.length;

          if (failedRecords.length > 0) {
            console.error('Failed records:', failedRecords);
            if (savedCount > 0) {
              alert(`Imported ${savedCount} of ${total} report(s). ${failedRecords.length} failed. Check console for details.`);
            } else {
              alert(`Import failed. Error: ${failedRecords[0]?.error || 'Unknown error'}. Check browser console (F12) for details.`);
            }
          } else {
            alert(`Imported ${total} report(s) successfully.`);
          }
          evt.target.value = '';
        } catch (err) {
          console.error('Bulk import error', err);
          alert('Import failed: ' + (err?.message || err || 'Unknown error'));
          evt.target.value = '';
        }
      };
      reader.onerror = () => { alert('Failed to read the uploaded file.'); evt.target.value = ''; };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      alert('Import failed: ' + (err?.message || err));
      evt.target.value = '';
    }
  }

  async function handleSurPdfImport(evt) {
    const file = evt.target.files[0];
    if (!file) return;
    try {
      showLoading(true);
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument(new Uint8Array(arrayBuffer)).promise;
      let fullTextLines = [];

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        const linesObj = {};
        textContent.items.forEach(item => {
          const x = item.transform[4];
          const y = Math.round(item.transform[5] / 2) * 2;
          if (!linesObj[y]) linesObj[y] = [];
          linesObj[y].push({ x, str: item.str });
        });
        
        const sortedY = Object.keys(linesObj).map(Number).sort((a,b) => b - a);
        sortedY.forEach(y => {
          const lineItems = linesObj[y].sort((a,b) => a.x - b.x);
          const lineStr = lineItems.map(item => item.str).join(' ').replace(/\s+/g, ' ').trim();
          if (lineStr) fullTextLines.push(lineStr);
        });
      }

      let reportYearStr = new Date().getFullYear().toString();
      for (let i = 0; i < Math.min(20, fullTextLines.length); i++) {
        const line = fullTextLines[i];
        if (line.includes("WATER SERVICE UPDATE REPORT")) {
           let dateLine = fullTextLines[i+1] || "";
           const d = new Date(dateLine);
           if (!isNaN(d.getTime())) {
             reportYearStr = d.getFullYear().toString();
           }
        }
      }

      const monthMap = { 'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 
                         'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12' };

      const parseDatesFromLine = (line, prefix) => {
          const str = line.substring(prefix.length).trim();
          const matches = str.split(/\s+/);
          return matches.map(rd => {
              const parts = rd.split('-');
              if (parts.length === 2) {
                  const dd = parts[0].padStart(2, '0');
                  const mm = monthMap[parts[1]] || '01';
                  return `${reportYearStr}-${mm}-${dd}`;
              }
              return null;
          });
      };

      const getLineValues = (line, prefix) => {
          const str = line.substring(prefix.length).trim();
          return str.split(/\s+/);
      };

      let damDates = [];
      let inflowDates = [];
      let prodDates = [];
      let augDates = [];

      fullTextLines.forEach(line => {
          if (line.startsWith("DAM LEVELS (as of 6AM)")) damDates = parseDatesFromLine(line, "DAM LEVELS (as of 6AM)");
          if (line.startsWith("RAW WATER INFLOWS")) inflowDates = parseDatesFromLine(line, "RAW WATER INFLOWS (MLD)");
          if (line.startsWith("DAILY PRODUCTION")) prodDates = parseDatesFromLine(line, "DAILY PRODUCTION (MLD)");
          if (line.startsWith("SUPPLY AUGMENTATION")) augDates = parseDatesFromLine(line, "SUPPLY AUGMENTATION (MLD)");
      });

      const recordsByDate = {};
      const getRecord = (dateStr) => {
          if (!dateStr) return null;
          if (!recordsByDate[dateStr]) {
              const existing = serviceUpdates.find(x => (x.date || '') === dateStr && ((x.provider || '').toUpperCase() === "MWCI"));
              recordsByDate[dateStr] = {
                  dateStr,
                  angat: existing?.angat || 0,
                  ipo: existing?.ipo || 0,
                  laMesa: existing?.laMesa || 0,
                  deepwell: 0, crossBorder: 0,
                  rawInflowsMap: {}, productionMap: {},
                  hasDam: !!(existing?.angat || existing?.ipo || existing?.laMesa),
                  hasPlants: !!(existing?.plants && existing.plants.length > 0),
                  hasAug: !!(existing?.supplyAug),
                  existingSupplyAug: existing?.supplyAug || 0
              };
              // Pre-fill plant maps if existing data exists
              if (existing && Array.isArray(existing.plants)) {
                  existing.plants.forEach(p => {
                      recordsByDate[dateStr].rawInflowsMap[p.name] = p.inflow || 0;
                      recordsByDate[dateStr].productionMap[p.name] = p.production || 0;
                  });
              }
          }
          return recordsByDate[dateStr];
      };

      const fillMetric = (prefix, datesArr, prop, flag) => {
          const line = fullTextLines.find(l => l.startsWith(prefix));
          if (line) {
              const vals = getLineValues(line, prefix);
              for (let i = 0; i < Math.min(datesArr.length, vals.length); i++) {
                  const dateStr = datesArr[i];
                  if (dateStr) {
                      const rec = getRecord(dateStr);
                      const num = parseFloat(vals[i].replace(/,/g, ''));
                      const val = isNaN(num) ? 0 : num;
                      // Only update if parsed value is non-zero, or if existing value is zero
                      if (val > 0 || (rec[prop] || 0) === 0) {
                          rec[prop] = val;
                          if (flag) rec[flag] = true;
                      }
                  }
              }
          }
      };

      fillMetric("Angat Dam (m)", damDates, "angat", "hasDam");
      fillMetric("Ipo Dam (m)", damDates, "ipo", "hasDam");
      fillMetric("La Mesa Dam (m)", damDates, "laMesa", "hasDam");

      fillMetric("Deepwell Production", augDates, "deepwell", "hasAug");
      fillMetric("Cross Border Flow", augDates, "crossBorder", "hasAug");

      const plantMap = [
          { pdfName: "Balara WTP 1", elName: "Balara WTP 1" },
          { pdfName: "Balara WTP 2", elName: "Balara WTP 2" },
          { pdfName: "East La Mesa WTP", elName: "East LMTP" },
          { pdfName: "Luzon WTP", elName: "Luzon WTP" },
          { pdfName: "Cardona WTP", elName: "Cardona WTP" },
          { pdfName: "Calawis WTP", elName: "Calawis TP" },
          { pdfName: "East Bay Phase 1 WTP", elName: "East Bay TP" }
      ];

      let currentSectionConfig = { prefix: "", dates: [] };

      fullTextLines.forEach(line => {
          if (line.startsWith("RAW WATER INFLOWS")) { currentSectionConfig = { prefix: "inflow", dates: inflowDates }; return; }
          if (line.startsWith("DAILY PRODUCTION")) { currentSectionConfig = { prefix: "prod", dates: prodDates }; return; }
          if (line.startsWith("SUPPLY AUGMENTATION")) { currentSectionConfig = { prefix: "", dates: [] }; return; }
          if (line.startsWith("DAM LEVELS")) { currentSectionConfig = { prefix: "", dates: [] }; return; }
          
          if (currentSectionConfig.prefix) {
              for (let p of plantMap) {
                  if (line.startsWith(p.pdfName)) {
                      const vals = getLineValues(line, p.pdfName);
                      for (let i = 0; i < Math.min(currentSectionConfig.dates.length, vals.length); i++) {
                          const dateStr = currentSectionConfig.dates[i];
                          if (dateStr) {
                              const rec = getRecord(dateStr);
                              const num = parseFloat(vals[i].replace(/,/g, ''));
                              const val = isNaN(num) ? 0 : num;
                              if (currentSectionConfig.prefix === "inflow") {
                                  if (val > 0 || (rec.rawInflowsMap[p.elName] || 0) === 0) {
                                      rec.rawInflowsMap[p.elName] = val;
                                  }
                              }
                              if (currentSectionConfig.prefix === "prod") {
                                  if (val > 0 || (rec.productionMap[p.elName] || 0) === 0) {
                                      rec.productionMap[p.elName] = val;
                                  }
                              }
                              rec.hasPlants = true;
                          }
                      }
                  }
              }
          }
      });

      const toSave = [];
      const timestamp = Date.now();
      
      const allDates = Object.keys(recordsByDate).sort();
      for (const dStr of allDates) {
          const d = recordsByDate[dStr];
          const calculatedAug = Math.round((d.deepwell + d.crossBorder) * 100) / 100;
          // Use PDF values if they result in > 0, otherwise keep existing
          d.supplyAug = (calculatedAug > 0 || !d.existingSupplyAug) ? calculatedAug : d.existingSupplyAug;
          
          d.plants = [];
          
          for (let p of plantMap) {
              const raw = d.rawInflowsMap[p.elName] || 0;
              const prod = d.productionMap[p.elName] || 0;
              d.plants.push({ name: p.elName, inflow: raw, production: prod });
          }
          const totalInflows = Math.round(d.plants.reduce((sum, p) => sum + p.inflow, 0) * 100) / 100;
          const totalProd = Math.round(d.plants.reduce((sum, p) => sum + p.production, 0) * 100) / 100;
          
          const s = {
             date: d.dateStr,
             provider: "MWCI",
             damAsOf: "06:00",
             inflows: totalInflows,
             production: totalProd,
             supplyAug: d.supplyAug || 0,
             angat: d.angat || 0,
             ipo: d.ipo || 0,
             laMesa: d.laMesa || 0,
             plants: d.plants || [],
             plantsDate: d.dateStr,
             remarks: "Imported via PDF",
             timestamp: timestamp
          };

          const existing = serviceUpdates.find(x => (x.date || '') === d.dateStr && ((x.provider || '').toUpperCase() === "MWCI"));
          if (existing) { s.id = existing.id; }
          toSave.push(s);
      }

      if (toSave.length === 0) {
        alert("Found no data rows to import.");
        return;
      }

      const confirmMsg = `Import ${toSave.length} Service Update(s) from PDF spanning ${toSave[0].date} to ${toSave[toSave.length-1].date}? This will create/update records.`;
      if (!confirm(confirmMsg)) return;

      // Save in parallel
      const results = await Promise.all(toSave.map(async (item) => {
        try {
          await saveServiceUpdate(item);
          return { success: true };
        } catch (saveErr) {
          console.error(`Failed to save record for ${item.date}:`, saveErr);
          return { success: false, date: item.date, error: saveErr?.message || String(saveErr) };
        }
      }));

      const failedRecords = results.filter(r => !r.success);
      const savedCount = results.length - failedRecords.length;

      if (failedRecords.length > 0) {
        alert(`Imported ${savedCount} of ${toSave.length} report(s). ${failedRecords.length} failed. Check console for details.`);
      } else {
        alert(`Successfully extracted and imported data for ${toSave.length} dates!`);
      }

    } catch (err) {
      console.error('PDF Import error', err);
      alert('PDF Import Extract failed: ' + (err?.message || err || 'Unknown error'));
    } finally {
      evt.target.value = '';
      showLoading(false);
    }
  }

  /**
   * MWSI (Maynilad) PDF Import handler
   * Parses the Maynilad Water Service Update Report PDF.
   * - Maps plant names from PDF to system plant names
   * - Cross-border flow items (Aurora Blvd, San Martin de Porres, BF Mapayapa, A. Francisco, East La Mesa) are summed into "Cross-Border Flow"
   * - Deepwell row in Daily Production is mapped to Supply Augmentation
   */
  async function handleSurPdfImportMwsi(evt) {
    const file = evt.target.files[0];
    if (!file) return;
    try {
      showLoading(true);
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument(new Uint8Array(arrayBuffer)).promise;
      let fullTextLines = [];

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const linesObj = {};
        textContent.items.forEach(item => {
          const x = item.transform[4];
          const y = Math.round(item.transform[5] / 2) * 2;
          if (!linesObj[y]) linesObj[y] = [];
          linesObj[y].push({ x, str: item.str });
        });
        const sortedY = Object.keys(linesObj).map(Number).sort((a, b) => b - a);
        sortedY.forEach(y => {
          const lineItems = linesObj[y].sort((a, b) => a.x - b.x);
          const lineStr = lineItems.map(item => item.str).join(' ').replace(/\s+/g, ' ').trim();
          if (lineStr) fullTextLines.push(lineStr);
        });
      }

      // Detect report year from header
      let reportYearStr = new Date().getFullYear().toString();
      for (let i = 0; i < Math.min(20, fullTextLines.length); i++) {
        const line = fullTextLines[i];
        if (line.includes('WATER SERVICE UPDATE REPORT') || line.startsWith('Date:')) {
          const candidates = [line, fullTextLines[i + 1] || ''];
          for (const c of candidates) {
            const cleaned = c.replace(/Date:\s*/i, '').replace(/(\d)\s+(\d)/g, '$1$2').replace(/\s*,/g, ',').trim();
            const d = new Date(cleaned);
            if (!isNaN(d.getTime())) {
              reportYearStr = d.getFullYear().toString();
              break;
            }
          }
        }
      }

      const monthMap = { 'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
                         'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12' };

      // Parse date headers like "23 - Mar 24 - Mar 25 - Mar ..."
      let columnDates = [];
      for (let i = 0; i < fullTextLines.length; i++) {
        // Pre-clean: collapse spaces between adjacent digits (e.g. "2 7" -> "27")
        const line = fullTextLines[i].replace(/(\d)\s+(\d)/g, '$1$2');
        const datePattern = /(\d{1,2})\s*-\s*(\w{3})/g;
        let match;
        const found = [];
        while ((match = datePattern.exec(line)) !== null) {
          found.push({ day: match[1], mon: match[2] });
        }
        if (found.length >= 3) {
          columnDates = found.map(f => {
            const dd = f.day.padStart(2, '0');
            const mm = monthMap[f.mon] || '01';
            return `${reportYearStr}-${mm}-${dd}`;
          });
          break;
        }
      }

      if (columnDates.length === 0) {
        alert('Could not detect date columns in the MWSI PDF. Please check the file format.');
        return;
      }

      // Plant mapping: PDF name prefix -> system name
      const mwsiPlantMap = [
        { pdfPrefix: 'La Mesa Water Treatment Plant 1', elName: 'La Mesa WTP 1' },
        { pdfPrefix: 'La Me sa Water Treatment Plant 1', elName: 'La Mesa WTP 1' },
        { pdfPrefix: 'La Mesa Water Treatment Plant 2', elName: 'La Mesa WTP 2' },
        { pdfPrefix: 'La Me sa Water Treatment Plant 2', elName: 'La Mesa WTP 2' },
        { pdfPrefix: 'Putatan 1', elName: 'Putatan WTP 1' },
        { pdfPrefix: 'Putatan 2', elName: 'Putatan WTP 2' },
        { pdfPrefix: 'PQE New Water', elName: 'PQE New Water' },
        { pdfPrefix: 'Anabu', elName: 'Anabu WTP' },
        { pdfPrefix: 'Poblacion WTP', elName: 'Poblacion WTP' },
        { pdfPrefix: 'Poblacion', elName: 'Poblacion WTP' },
        { pdfPrefix: 'Nanostone', elName: 'Nanostone' },
      ];

      // Cross-border flow items to sum
      const crossBorderPrefixes = [
        'Aurora Blvd', 'Aurora blvd',
        'San Martin de Porres', 'San Martin De Porres',
        'BF Mapayapa',
        'A. Francisco',
        'East La Mesa',
      ];

      // Build per-date records
      const recordsByDate = {};
      const getRecord = (dateStr) => {
        if (!dateStr) return null;
        if (!recordsByDate[dateStr]) {
          const existing = serviceUpdates.find(x => (x.date || '') === dateStr && ((x.provider || '').toUpperCase() === 'MWSI'));
          recordsByDate[dateStr] = {
            dateStr,
            angat: existing?.angat || 0,
            ipo: existing?.ipo || 0,
            laMesa: existing?.laMesa || 0,
            supplyAug: existing?.supplyAug || 0,
            rawInflowsMap: {},
            productionMap: {},
            crossBorderProdByDate: 0,
            deepwellByDate: 0,
            hasDam: !!(existing?.angat || existing?.ipo || existing?.laMesa),
            hasPlants: !!(existing?.plants && existing.plants.length > 0),
            existingSupplyAug: existing?.supplyAug || 0
          };
          if (existing && Array.isArray(existing.plants)) {
            existing.plants.forEach(p => {
              recordsByDate[dateStr].rawInflowsMap[p.name] = p.inflow || 0;
              recordsByDate[dateStr].productionMap[p.name] = p.production || 0;
            });
          }
        }
        return recordsByDate[dateStr];
      };

      // Helper to extract numeric values from a line after its prefix
      const extractValues = (line, prefix) => {
        let rest = line.substring(prefix.length).trim();
        // Pre-clean PDF artifacts: collapse split numbers like "8 .02" -> "8.02", "17. 55" -> "17.55"
        rest = rest.replace(/(\d)\s+\./g, '$1.').replace(/\.\s+(\d)/g, '.$1');
        
        const tokens = rest.split(/\s+/);
        const parsedValues = [];
        for (const t of tokens) {
          if (t === '*') continue; // ignore detached single asterisks (annotations)
          
          const cleaned = t.replace(/[*]/g, '').replace(/,/g, '').trim();
          if (!cleaned) {
            // Either "**" missing value placeholder or just something we completely erased
            if (t.includes('**') || t === '-' || t === '') {
              parsedValues.push(null);
            } else {
              parsedValues.push(null);
            }
          } else {
            const num = parseFloat(cleaned);
            parsedValues.push(isNaN(num) ? null : num);
          }
        }
        return parsedValues;
      };

      // Parse sections
      let currentSection = '';

      fullTextLines.forEach(line => {
        if (line.startsWith('RAW WATER INFLOWS')) { currentSection = 'inflow'; return; }
        if (line.startsWith('DAILY PRODUCTION')) { currentSection = 'prod'; return; }
        if (line.startsWith('CROSS') && line.includes('BORDER')) { currentSection = 'crossborder'; return; }
        if (line.startsWith('WATER AVAILABILITY') || line.startsWith('WATER SERVICE INTERRUPTIONS') || line.startsWith('SUPPLY (%)') || line.startsWith('BUSINESS AREA')) {
          currentSection = '';
          return;
        }
        if (line.startsWith('Note:') || (line.startsWith('*') && !line.startsWith('*D') === false)) return;
        if (line.startsWith('Note:')) return;

        if (currentSection === 'inflow' || currentSection === 'prod') {
          // Deepwell row in production -> Supply Augmentation
          if (currentSection === 'prod' && line.startsWith('Deepwell')) {
            const vals = extractValues(line, 'Deepwell');
            for (let i = 0; i < Math.min(columnDates.length, vals.length); i++) {
              if (columnDates[i] && vals[i] !== null) {
                const rec = getRecord(columnDates[i]);
                if (rec) rec.deepwellByDate = vals[i];
              }
            }
            return;
          }

          // Match against plant map (longest prefix first)
          let matched = false;
          const sortedPlantMap = [...mwsiPlantMap].sort((a, b) => b.pdfPrefix.length - a.pdfPrefix.length);
          for (const p of sortedPlantMap) {
            if (line.startsWith(p.pdfPrefix)) {
              const vals = extractValues(line, p.pdfPrefix);
              for (let i = 0; i < Math.min(columnDates.length, vals.length); i++) {
                if (columnDates[i] && vals[i] !== null) {
                  const rec = getRecord(columnDates[i]);
                  if (currentSection === 'inflow') {
                    if (vals[i] > 0 || (rec.rawInflowsMap[p.elName] || 0) === 0) {
                      rec.rawInflowsMap[p.elName] = vals[i];
                    }
                  } else {
                    if (vals[i] > 0 || (rec.productionMap[p.elName] || 0) === 0) {
                      rec.productionMap[p.elName] = vals[i];
                    }
                  }
                  rec.hasPlants = true;
                }
              }
              matched = true;
              break;
            }
          }
          if (matched) return;
        }

        if (currentSection === 'crossborder') {
          for (const cbPrefix of crossBorderPrefixes) {
            if (line.startsWith(cbPrefix)) {
              const vals = extractValues(line, cbPrefix);
              for (let i = 0; i < Math.min(columnDates.length, vals.length); i++) {
                if (columnDates[i] && vals[i] !== null) {
                  const rec = getRecord(columnDates[i]);
                  rec.crossBorderProdByDate = (rec.crossBorderProdByDate || 0) + vals[i];
                }
              }
              return;
            }
          }
        }
      });

      // Build final records
      const toSave = [];
      const timestamp = Date.now();
      const systemPlants = MWSI_PLANTS;

      const allDates = Object.keys(recordsByDate).sort();
      for (const dStr of allDates) {
        const d = recordsByDate[dStr];

        // Set cross-border flow production value
        if (d.crossBorderProdByDate > 0 || (d.productionMap['Cross-Border Flow'] || 0) === 0) {
          d.productionMap['Cross-Border Flow'] = Math.round(d.crossBorderProdByDate * 100) / 100;
        }

        // Deepwell -> Supply Augmentation
        const supplyAug = d.deepwellByDate > 0 ? d.deepwellByDate : d.existingSupplyAug;

        // Build plants array
        const plants = [];
        for (const plantName of systemPlants) {
          plants.push({
            name: plantName,
            inflow: d.rawInflowsMap[plantName] || 0,
            production: d.productionMap[plantName] || 0
          });
        }

        const totalInflows = Math.round(plants.reduce((sum, p) => sum + p.inflow, 0) * 100) / 100;
        const totalProd = Math.round(plants.reduce((sum, p) => sum + p.production, 0) * 100) / 100;

        const s = {
          date: dStr,
          provider: 'MWSI',
          damAsOf: '',
          inflows: totalInflows,
          production: totalProd,
          supplyAug: supplyAug || 0,
          angat: d.angat || 0,
          ipo: d.ipo || 0,
          laMesa: d.laMesa || 0,
          plants: plants,
          plantsDate: dStr,
          remarks: 'Imported via PDF (MWSI)',
          timestamp: timestamp
        };

        const existing = serviceUpdates.find(x => (x.date || '') === dStr && ((x.provider || '').toUpperCase() === 'MWSI'));
        if (existing) { s.id = existing.id; }
        toSave.push(s);
      }

      if (toSave.length === 0) {
        alert('Found no data rows to import from the MWSI PDF.');
        return;
      }

      const confirmMsg = `Import ${toSave.length} MWSI Service Update(s) from PDF spanning ${toSave[0].date} to ${toSave[toSave.length - 1].date}? This will create/update records.`;
      if (!confirm(confirmMsg)) return;

      // Save in parallel
      const results = await Promise.all(toSave.map(async (item) => {
        try {
          await saveServiceUpdate(item);
          return { success: true };
        } catch (saveErr) {
          console.error(`Failed to save MWSI record for ${item.date}:`, saveErr);
          return { success: false, date: item.date, error: saveErr?.message || String(saveErr) };
        }
      }));

      const failedRecords = results.filter(r => !r.success);
      const savedCount = results.length - failedRecords.length;

      if (failedRecords.length > 0) {
        alert(`Imported ${savedCount} of ${toSave.length} MWSI report(s). ${failedRecords.length} failed. Check console for details.`);
      } else {
        alert(`Successfully extracted and imported MWSI data for ${toSave.length} dates!`);
      }

    } catch (err) {
      console.error('MWSI PDF Import error', err);
      alert('MWSI PDF Import failed: ' + (err?.message || err || 'Unknown error'));
    } finally {
      evt.target.value = '';
      showLoading(false);
    }
  }

  function getFilteredServiceUpdates() {
    let list = serviceUpdates.slice();
    const start = elements.serviceStartDateFilter?.value || '';
    const end = elements.serviceEndDateFilter?.value || '';
    const prov = elements.serviceProviderFilter?.value || '';
    const provCanon = prov ? canonProvider(prov) : '';
    if (start) list = list.filter(s => safeDateCompare((s.date || ''), start) >= 0);
    if (end) list = list.filter(s => safeDateCompare((s.date || ''), end) <= 0);
    if (provCanon) list = list.filter(s => canonProvider(s.provider || '') === provCanon);
    // sort same as table
    list.sort((a, b) => {
      const ad = a.date || ''; const bd = b.date || '';
      const cmp = safeDateCompare(ad, bd);
      if (cmp !== 0) return cmp > 0 ? -1 : 1; // desc by date
      return (a.provider || '').localeCompare(b.provider || '');
    });
    return list;
  }

  function exportSurReports() {
    try {
      const list = getFilteredServiceUpdates();
      if (list.length === 0) { alert('No reports to export for the selected filters.'); return; }
      // Build columns
      const baseHeaders = [
        'Date', 'Concessionaire', 'Dam As of', 'Supply Augmentation (MLD)', 'Angat Dam (masl)', 'Ipo Dam (masl)', 'La Mesa Dam (masl)', 'Total Raw Inflows (MLD)', 'Total Treatment Production (MLD)'
      ];
      const unionPlants = [...new Set([...getProviderPlants('MWCI'), ...getProviderPlants('MWSI')])];
      const inflowHeaders = [];
      const productionHeaders = [];
      unionPlants.forEach(p => inflowHeaders.push(`${p} - Inflow (MLD)`));
      unionPlants.forEach(p => productionHeaders.push(`${p} - Production (MLD)`));
      const headers = [...baseHeaders, ...inflowHeaders, ...productionHeaders];
      const rows = list.map(s => {
        const r = {
          'Date': normalizeYmd(s.date || ''),
          'Concessionaire': s.provider || '',
          'Dam As of': s.damAsOf || '',
          'Supply Augmentation (MLD)': Number(s.supplyAug || 0),
          'Angat Dam (masl)': Number(s.angat || 0),
          'Ipo Dam (masl)': Number(s.ipo || 0),
          'La Mesa Dam (masl)': Number(s.laMesa || 0),
          'Total Raw Inflows (MLD)': Number(s.inflows || 0),
          'Total Treatment Production (MLD)': Number(s.production || 0)
        };
        const byName = new Map();
        (Array.isArray(s.plants) ? s.plants : []).forEach(p => { byName.set((p.name || '').toString(), p); });
        unionPlants.forEach(pn => {
          const v = byName.get(pn) || {};
          r[`${pn} - Inflow (MLD)`] = Number(v.inflow || 0);
          r[`${pn} - Production (MLD)`] = Number(v.production || 0);
        });
        return r;
      });
      const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Reports');
      const start = elements.serviceStartDateFilter?.value || '';
      const end = elements.serviceEndDateFilter?.value || '';
      const prov = elements.serviceProviderFilter?.value || 'all';
      const fname = `SUR_Export_${prov}_${start || 'all'}_${end || 'all'}.xlsx`;
      XLSX.writeFile(wb, fname);
    } catch (err) {
      console.error('Export error', err);
      alert('Failed to export reports.');
    }
  }

  // --- Read-only details rendering for Service Update Report ---
  function renderSurPlantDetailsTable(plants) {
    const container = document.getElementById('surDetPlantTableContainer');
    if (!container) return;
    const list = Array.isArray(plants) ? plants : [];
    if (list.length === 0) {
      container.innerHTML = '<div class="text-muted small">No per-plant metrics recorded.</div>';
      return;
    }
    function fmt2(n) { return (Math.round((n || 0) * 100) / 100).toFixed(2); }
    let html = '';
    html += '<div class="table-responsive">';
    html += '<table class="table table-sm align-middle mb-2">';
    html += '<thead><tr class="table-light"><th>Plant</th><th style="width:160px">Raw Inflows (MLD)</th><th style="width:200px">Daily Production (MLD)</th></tr></thead><tbody>';
    list.forEach(p => {
      html += '<tr>'
        + `<td class="small">${p.name || ''}</td>`
        + `<td>${fmt2(p.inflow)}</td>`
        + `<td>${fmt2(p.production)}</td>`
        + '</tr>';
    });
    html += '</tbody></table>';
    html += '</div>';
    container.innerHTML = html;
  }

  function viewServiceUpdateDetails(id) {
    const s = serviceUpdates.find(x => x.id === id);
    if (!s) return;
    // Fill basic fields
    const el = (id) => document.getElementById(id);
    if (el('surDetDate')) el('surDetDate').textContent = s.date || '';
    if (el('surDetProvider')) el('surDetProvider').textContent = s.provider || '';
    if (el('surDetDamDate')) el('surDetDamDate').textContent = s.date || '';
    if (el('surDetDamAsOf')) el('surDetDamAsOf').textContent = s.damAsOf ? `(${s.damAsOf})` : '';
    if (el('surDetPlantsDate')) el('surDetPlantsDate').textContent = s.date || '';
    // Plants table
    const plantsData = Array.isArray(s.plants) ? s.plants : [];
    renderSurPlantDetailsTable(plantsData);
    // Totals and dams
    const sumIn = (Array.isArray(plantsData) ? plantsData : []).reduce((a, p) => a + (p.inflow || 0), 0);
    const sumPr = (Array.isArray(plantsData) ? plantsData : []).reduce((a, p) => a + (p.production || 0), 0);
    const inflows = (typeof s.inflows === 'number') ? s.inflows : sumIn;
    const production = (typeof s.production === 'number') ? s.production : sumPr;
    const fmt2 = (n) => (Math.round((n || 0) * 100) / 100).toFixed(2);
    if (el('surDetInflows')) el('surDetInflows').textContent = fmt2(inflows);
    if (el('surDetProduction')) el('surDetProduction').textContent = fmt2(production);
    if (el('surDetSupplyAug')) el('surDetSupplyAug').textContent = fmt2(s.supplyAug || 0);
    if (el('surDetAngat')) el('surDetAngat').textContent = fmt2(s.angat || 0);
    if (el('surDetIpo')) el('surDetIpo').textContent = fmt2(s.ipo || 0);
    if (el('surDetLaMesa')) el('surDetLaMesa').textContent = fmt2(s.laMesa || 0);
    if (el('surDetRemarks')) el('surDetRemarks').textContent = s.remarks || '';
    // History
    const histEl = el('surDetHistory');
    if (histEl) {
      const hist = Array.isArray(s.history) ? [...s.history] : [];
      hist.sort((a, b) => {
        const at = a.timestamp || '';
        const bt = b.timestamp || '';
        return (at > bt ? -1 : at < bt ? 1 : 0);
      });
      if (hist.length === 0) {
        histEl.innerHTML = '<span class="text-muted">No edits yet.</span>';
      } else {
        const rows = hist.map(h => {
          const when = h.timestamp ? new Date(h.timestamp).toLocaleString() : '';
          const who = h.email || 'Unknown user';
          const act = h.action || 'update';
          return `<div class="d-flex align-items-start gap-2 py-1">
          <i class="fa-regular fa-clock mt-1 text-muted"></i>
          <div>
            <div><strong>${act}</strong> by ${who}</div>
            <div class="text-muted small">${when}</div>
          </div>
        </div>`;
        }).join('');
        histEl.innerHTML = rows;
      }
    }
    // Show read-only modal
    if (elements.serviceUpdateDetailsModal) elements.serviceUpdateDetailsModal.show();
  }
  // Expose globally for row click handlers
  window.viewServiceUpdateDetails = viewServiceUpdateDetails;

  function serviceUpdateRowHtml(s) {
    let actionsHtml = elevatedAccess ? `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="editServiceUpdate('${s.id}', true)"><i class="fa fa-pencil"></i></button>` : '';
    if (isAdmin) {
      actionsHtml += `<button class="btn btn-sm btn-danger" title="Delete" onclick="deleteServiceUpdate('${s.id}')"><i class="fa fa-trash"></i></button>`;
    }
    return `<tr data-id="${s.id}">`
      + `<td>${s.date || ''}</td>`
      + `<td>${s.provider || ''}</td>`
      + `<td>${(Array.isArray(s.plants) ? s.plants.length : 0)}</td>`
      + `<td>${fmtNum(s.inflows || 0)}</td>`
      + `<td>${fmtNum(s.production || 0)}</td>`
      + `<td>${fmtNum(s.supplyAug || 0)}</td>`
      + `<td>${fmtNum(s.angat || 0)}</td>`
      + `<td>${fmtNum(s.ipo || 0)}</td>`
      + `<td>${fmtNum(s.laMesa || 0)}</td>`
      + `<td>${s.damAsOf || ''}</td>`
      + `<td class="d-flex gap-1">${actionsHtml}</td>`
      + `</tr>`;
  }

  // Legacy renderer (kept for reference; main rendering delegated to ServiceUpdatesFeatureInstance)
  function renderServiceUpdatesLegacy() {
    if (elements.addServiceUpdateBtn) {
      elements.addServiceUpdateBtn.style.display = (elevatedAccess && auth.currentUser && !isViewOnly) ? 'inline-block' : 'none';
    }
    // Toggle bulk action buttons visibility
    const bulkImpBtn = document.getElementById('surBulkImportBtn');
    if (bulkImpBtn) {
      bulkImpBtn.style.display = (elevatedAccess && auth.currentUser && !isViewOnly) ? 'inline-block' : 'none';
    }
    const bulkTplBtn = document.getElementById('surBulkTemplateBtn');
    if (bulkTplBtn) { bulkTplBtn.style.display = 'inline-block'; }
    const exportBtn = document.getElementById('surExportBtn');
    if (exportBtn) { exportBtn.style.display = 'inline-block'; }
    let list = serviceUpdates.slice();
    const start = elements.serviceStartDateFilter?.value || '';
    const end = elements.serviceEndDateFilter?.value || '';
    const prov = elements.serviceProviderFilter?.value || '';
    if (start) list = list.filter(s => (s.date || '') >= start);
    if (end) list = list.filter(s => (s.date || '') <= end);
    if (prov) list = list.filter(s => (s.provider || '') === prov);
    // sort by date desc, then provider
    list.sort((a, b) => {
      const ad = a.date || ''; const bd = b.date || '';
      const cmp = safeDateCompare(ad, bd);
      if (cmp !== 0) return cmp > 0 ? -1 : 1;
      return (a.provider || '').localeCompare(b.provider || '');
    });
    // Pagination
    const total = list.length;
    const perPage = SERVICE_UPDATES_PER_PAGE;
    const maxPage = Math.max(1, Math.ceil(total / perPage));
    if (serviceUpdatesPage > maxPage) serviceUpdatesPage = maxPage;
    const startIdx = (serviceUpdatesPage - 1) * perPage;
    const endIdx = Math.min(startIdx + perPage, total);
    const pageSlice = list.slice(startIdx, endIdx);
    const tbody = elements.serviceUpdatesTbody;
    if (!tbody) return;
    tbody.innerHTML = pageSlice.map(serviceUpdateRowHtml).join('');
    // row click to edit (excluding buttons)
    Array.from(tbody.querySelectorAll('tr')).forEach(tr => {
      tr.style.cursor = 'pointer';
      tr.addEventListener('click', e => {
        if (e.target.closest('button')) return;
        // Open read-only details modal on row click
        viewServiceUpdateDetails(tr.dataset.id);
      });
    });
    // Update pagination UI
    const infoEl = elements.serviceUpdatesPageInfo;
    if (infoEl) {
      if (total === 0) { infoEl.textContent = '0â€“0 of 0'; }
      else { infoEl.textContent = `${(startIdx + 1).toLocaleString()}â€“${endIdx.toLocaleString()} of ${total.toLocaleString()}`; }
    }
    if (elements.serviceUpdatesPrev) { elements.serviceUpdatesPrev.disabled = (serviceUpdatesPage <= 1); }
    if (elements.serviceUpdatesNext) { elements.serviceUpdatesNext.disabled = (serviceUpdatesPage >= maxPage); }
  }

  function clearServiceUpdateForm() {
    const f = elements.serviceUpdateForm;
    if (!f) return;
    f.reset();
    if (f.surId) f.surId.value = '';
    // render empty per-plant table for current provider
    const prov = f.surProvider?.value || 'MWCI';
    renderSurPlantTable(prov, []);
    updateSurDateLabels();
    // Try to prefill dam elevations from counterpart if available
    prefillDamFromCounterpart();
  }

  async function prefillDamFromCounterpart() {
    const f = elements.serviceUpdateForm;
    if (!f) return;
    const date = f.surDate?.value || '';
    const provider = f.surProvider?.value || 'MWCI';
    if (!date) return;
    const otherProvider = provider === 'MWCI' ? 'MWSI' : 'MWCI';
    try {
      const snap = await serviceUpdateService.collection()
        .where('date', '==', date)
        .where('provider', '==', otherProvider)
        .limit(1)
        .get();
      if (snap.empty) return;
      const other = snap.docs[0].data();
      const otherHasDam = (Number(other.angat) > 0 || Number(other.ipo) > 0 || Number(other.laMesa) > 0);
      if (!otherHasDam) return;
      const hasCurrDam = (Number(f.surAngat?.value || 0) > 0 || Number(f.surIpo?.value || 0) > 0 || Number(f.surLaMesa?.value || 0) > 0);
      if (hasCurrDam) return; // don't override user-entered
      if (f.surAngat) f.surAngat.value = other.angat || 0;
      if (f.surIpo) f.surIpo.value = other.ipo || 0;
      if (f.surLaMesa) f.surLaMesa.value = other.laMesa || 0;
      if (f.surDamAsOf && !f.surDamAsOf.value) f.surDamAsOf.value = other.damAsOf || date;
    } catch (err) { console.warn('prefillDamFromCounterpart failed', err); }
  }

  async function saveServiceUpdate(s) {
    const userEmail = auth.currentUser?.email || 'viewer';
    // 1) Find counterpart doc for same date (other provider)
    const otherProvider = s.provider === 'MWCI' ? 'MWSI' : 'MWCI';
    let otherDoc = serviceUpdates.find(x => (x.date || '') === s.date && ((x.provider || '').toUpperCase() === otherProvider)) || null;
    if (!otherDoc) {
      try {
        const q = await serviceUpdateService.collection()
          .where('date', '==', s.date)
          .where('provider', '==', otherProvider)
          .limit(1)
          .get();
        if (!q.empty) {
          const d = q.docs[0];
          otherDoc = { id: d.id, ...d.data() };
        }
      } catch (err) { console.warn('Lookup counterpart failed', err); }
    }

    // 2) If creating new and current has no dam elevations but counterpart has, copy theirs before saving
    const sHasDam = (Number(s.angat) > 0 || Number(s.ipo) > 0 || Number(s.laMesa) > 0);
    const oHasDam = otherDoc && (Number(otherDoc.angat) > 0 || Number(otherDoc.ipo) > 0 || Number(otherDoc.laMesa) > 0);
    // Also apply on edit: if current payload lacks dam elevations but counterpart has, use counterpart values
    if (!sHasDam && oHasDam) {
      s.angat = otherDoc.angat || 0;
      s.ipo = otherDoc.ipo || 0;
      s.laMesa = otherDoc.laMesa || 0;
      if (!s.damAsOf && otherDoc.damAsOf) s.damAsOf = otherDoc.damAsOf;
    }

    // 3) Append edit history and save this document
    const actionLabel = s.id ? 'edit' : 'create';

    // IF creating new (no id), check if another record already exists for this exact date and provider
    if (!s.id) {
      const cachedExisting = serviceUpdates.find(x => (x.date || '') === s.date && (x.provider || '') === s.provider);
      if (cachedExisting) {
        throw new Error(`A report for ${s.provider} on ${s.date} already exists. Please edit the existing report instead.`);
      }
      const existingQuery = await serviceUpdateService.collection()
        .where('date', '==', s.date)
        .where('provider', '==', s.provider)
        .limit(1)
        .get();
      if (!existingQuery.empty) {
        throw new Error(`A report for ${s.provider} on ${s.date} already exists. Please edit the existing report instead.`);
      }
    }

    // Fetch existing history: use cached copy first, fallback to Firestore
    let existingHistory = [];
    if (s.id) {
      const cachedDoc = serviceUpdates.find(x => x.id === s.id);
      if (cachedDoc && Array.isArray(cachedDoc.history)) {
        existingHistory = cachedDoc.history;
      } else {
        try {
          const docSnap = await serviceUpdateService.collection().doc(s.id).get();
          existingHistory = docSnap.data()?.history || [];
        } catch (err) {
          console.warn('Failed to fetch existing history from DB', err);
        }
      }
    }

    const __userSUR = auth.currentUser;
    const userFullNameSUR = (__userSUR?.displayName || '').trim();
    s.history = [...existingHistory, { email: userEmail, fullName: userFullNameSUR, timestamp: new Date().toISOString(), action: actionLabel }];
    const data = { ...s };
    if (!data.id) delete data.id;
    let savedId = s.id;
    if (s.id) {
      await serviceUpdateService.collection().doc(s.id).set(data);
    } else {
      const ref = await serviceUpdateService.collection().add(data);
      savedId = ref.id;
    }

    // 4) Ensure counterpart matches current (keep same for same date), but don't overwrite non-zero with zeros
    if (otherDoc) {
      const mergedAngat = Number(s.angat) > 0 ? Number(s.angat) : Number(otherDoc.angat) || 0;
      const mergedIpo = Number(s.ipo) > 0 ? Number(s.ipo) : Number(otherDoc.ipo) || 0;
      const mergedLaMesa = Number(s.laMesa) > 0 ? Number(s.laMesa) : Number(otherDoc.laMesa) || 0;
      const mergedDamAsOf = (s.damAsOf || otherDoc.damAsOf || '');
      const needsSync = (
        Number(otherDoc.angat || 0) !== mergedAngat ||
        Number(otherDoc.ipo || 0) !== mergedIpo ||
        Number(otherDoc.laMesa || 0) !== mergedLaMesa ||
        ((otherDoc.damAsOf || '') !== mergedDamAsOf)
      );
      if (needsSync) {
        try {
          const otherHist = Array.isArray(otherDoc.history) ? otherDoc.history : [];
          await serviceUpdateService.collection().doc(otherDoc.id).set({
            angat: mergedAngat,
            ipo: mergedIpo,
            laMesa: mergedLaMesa,
            damAsOf: mergedDamAsOf,
            history: [...otherHist, { email: userEmail, timestamp: new Date().toISOString(), action: 'sync_dam' }]
          }, { merge: true });
        } catch (err) { console.warn('Sync counterpart failed', err); }
      }
    }
  }

  function gatherServiceUpdateForm() {
    const f = elements.serviceUpdateForm;
    const plants = readSurPlantData();
    const inflowTotal = plants.reduce((sum, p) => sum + (p.inflow || 0), 0);
    const prodTotal = plants.reduce((sum, p) => sum + (p.production || 0), 0);
    return {
      id: f.surId.value || undefined,
      date: f.surDate.value,
      provider: f.surProvider.value,
      damAsOf: f.surDamAsOf.value,
      inflows: Math.round(inflowTotal * 100) / 100,
      production: Math.round(prodTotal * 100) / 100,
      supplyAug: parseFloat(f.surSupplyAug.value) || 0,
      angat: parseFloat(f.surAngat.value) || 0,
      ipo: parseFloat(f.surIpo.value) || 0,
      laMesa: parseFloat(f.surLaMesa.value) || 0,
      plants,
      plantsDate: f.surDate.value,
      remarks: f.surRemarks.value,
      timestamp: Date.now()
    };
  }

  function onSaveServiceUpdate(e) {
    e.preventDefault();
    if (!elevatedAccess) { alert('Only admin/level2 can save'); return; }
    const saveBtn = document.getElementById('saveServiceUpdateBtn');
    if (saveBtn) saveBtn.disabled = true;
    const data = gatherServiceUpdateForm();
    // Close modal quickly for UX
    elements.serviceUpdateModal.hide();
    saveServiceUpdate(data)
      .catch(err => alert(err.message))
      .finally(() => { if (saveBtn) saveBtn.disabled = false; });
  }

  window.editServiceUpdate = (id, edit = true) => {
    const s = serviceUpdates.find(x => x.id === id);
    if (!s) return;
    const f = elements.serviceUpdateForm;
    f.surId.value = s.id;
    f.surDate.value = s.date || '';
    f.surProvider.value = s.provider || '';
    f.surDamAsOf.value = s.damAsOf || '';
    // Render plant table with existing data and compute totals
    const plantsData = Array.isArray(s.plants) ? s.plants : [];
    renderSurPlantTable(f.surProvider.value, plantsData);
    // If no per-plant data exists but totals are present, seed first row with totals to preserve values
    if ((!plantsData || plantsData.length === 0) && (s.inflows || s.production)) {
      const container = document.getElementById('surPlantTableContainer');
      const firstInflow = container?.querySelector('input[data-role="inflow"]');
      const firstProd = container?.querySelector('input[data-role="production"]');
      if (firstInflow) { firstInflow.value = s.inflows || 0; }
      if (firstProd) { firstProd.value = s.production || 0; }
      updateSurTotals();
    }
    updateSurDateLabels();
    f.surSupplyAug.value = s.supplyAug || '';
    f.surAngat.value = s.angat || '';
    f.surIpo.value = s.ipo || '';
    f.surLaMesa.value = s.laMesa || '';
    f.surRemarks.value = s.remarks || '';
    // Toggle readonly/disabled and Save visibility
    Array.from(f.elements).forEach(el => {
      if (el.tagName === 'INPUT' || el.tagName === 'SELECT' || el.tagName === 'TEXTAREA') {
        el.readOnly = !edit;
        el.disabled = !edit && (el.tagName === 'SELECT');
      }
    });
    const saveBtn = document.getElementById('saveServiceUpdateBtn');
    if (saveBtn) saveBtn.style.display = edit ? 'inline-block' : 'none';
    elements.serviceUpdateModal.show();
  };

  if (!window.deleteServiceUpdate) {
    window.deleteServiceUpdate = async (id) => {
      if (!isAdmin) { alert('Only admin can delete'); return; }
      if (!confirm('Delete this report?')) return;
      try {
        await serviceUpdateService.remove(id);
      } catch (err) { alert(err.message); }
    };
  }

  function subscribeServiceUpdates() {
    if (unsubscribeServiceUpdates) unsubscribeServiceUpdates();

    // Get selected year from dropdown (defaults to current year)
    const selectedYear = elements.serviceYearFilter?.value || new Date().getFullYear().toString();
    const startOfYear = `${selectedYear}-01-01`;
    const endOfYear = `${selectedYear}-12-31`;

    // Create a year-filtered query to reduce reads (from ~5000 to ~365)
    const query = serviceUpdateService.collection()
      .where('date', '>=', startOfYear)
      .where('date', '<=', endOfYear)
      .orderBy('date', 'desc');

    unsubscribeServiceUpdates = query.onSnapshot({ includeMetadataChanges: true }, snap => {
      const fromCache = !!(snap.metadata && snap.metadata.fromCache);
      serviceUpdates = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      ServiceUpdatesFeatureInstance.setServiceUpdates(serviceUpdates);
      // Always update the table immediately for responsiveness
      renderServiceUpdates();
      // Only re-render charts when data is confirmed from server to avoid stale graphs
      if (!fromCache) {
        if (typeof renderSurChart === 'function') try { renderSurChart(); } catch (_) {/*noop*/ }
        try { renderDashboard(); } catch (_) {/*noop*/ }
      }
    }, err => {
      console.error('Service Updates subscription error:', err);
      // If index not created, show helpful message
      if (err.message && err.message.includes('index')) {
        console.warn('Firestore composite index required. Click the link in the console above to create it.');
      }
    });
  }

  // Deepwells subscription: populate list and charts
  function subscribeDeepwells() {
    if (unsubscribeDeepwells) unsubscribeDeepwells();
    unsubscribeDeepwells = deepwellService.subscribe(snap => {
      const list = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      deepwells = list;
      DeepwellsFeatureInstance.setDeepwells(list);
      try { renderDeepwells(); } catch (_) {/*noop*/ }
      try { renderDeepwellMonthlyChart(); } catch (_) {/*noop*/ }
      try { renderDashboard(); } catch (_) {/*noop*/ }
    });
  }




  // Attach listeners
  if (elements.serviceUpdateForm) {
    elements.serviceUpdateForm.addEventListener('submit', onSaveServiceUpdate);
    const provEl = document.getElementById('surProvider');
    if (provEl) {
      provEl.addEventListener('change', () => {
        renderSurPlantTable(provEl.value, []);
        prefillDamFromCounterpart();
      });
    }
    const dateEl = document.getElementById('surDate');
    if (dateEl) { dateEl.addEventListener('change', () => { updateSurDateLabels(); prefillDamFromCounterpart(); }); }
    // Wire Import button/file input (desktop/mobile use same IDs)
    const importBtn = document.getElementById('surImportBtn');
    const importFile = document.getElementById('surImportFile');
    if (importBtn && importFile) {
      importBtn.addEventListener('click', () => importFile.click());
      importFile.addEventListener('change', handleSurImportFile);
    }
    const templateBtn = document.getElementById('surTemplateBtn');
    if (templateBtn) { templateBtn.addEventListener('click', downloadSurTemplate); }
    // Bulk controls in main UI (desktop)
    const bulkTplBtn = document.getElementById('surBulkTemplateBtn');
    const bulkImpBtn = document.getElementById('surBulkImportBtn');
    const bulkImpFile = document.getElementById('surBulkImportFile');
    const exportBtn = document.getElementById('surExportBtn');
    if (bulkTplBtn) { bulkTplBtn.addEventListener('click', downloadSurBulkTemplate); }
    if (bulkImpBtn && bulkImpFile) {
      bulkImpBtn.addEventListener('click', () => bulkImpFile.click());
      bulkImpFile.addEventListener('change', handleSurBulkImportFile);
    }
    
    const pdfImpBtn = document.getElementById('surPdfImportBtn');
    const pdfImpFile = document.getElementById('surPdfImportFile');
    if (pdfImpBtn && pdfImpFile) {
      pdfImpBtn.addEventListener('click', () => pdfImpFile.click());
      pdfImpFile.addEventListener('change', handleSurPdfImport);
    }

    const pdfImpMwsiBtn = document.getElementById('surPdfImportMwsiBtn');
    const pdfImpMwsiFile = document.getElementById('surPdfImportMwsiFile');
    if (pdfImpMwsiBtn && pdfImpMwsiFile) {
      pdfImpMwsiBtn.addEventListener('click', () => pdfImpMwsiFile.click());
      pdfImpMwsiFile.addEventListener('change', handleSurPdfImportMwsi);
    }
    

  }

  window.cleanUpDuplicateServiceUpdates = async function () {
    if (!isAdmin) { alert('Only admin can run cleanup'); return; }
    if (!confirm('This will scan all Service Update Reports and remove duplicates (same date & provider). Proceed?')) return;

    try {
      showLoading(true);
      console.log('[Cleanup] Starting Service Update Report cleanup...');
      const snap = await serviceUpdateService.collection().get();
      const allDocs = snap.docs.map(d => ({ id: d.id, ...d.data() }));

      const groups = {};
      allDocs.forEach(doc => {
        const key = `${doc.date}_${(doc.provider || '').toUpperCase()}`;
        if (!groups[key]) groups[key] = [];
        groups[key].push(doc);
      });

      const batch = db.batch();
      let removedCount = 0;

      Object.keys(groups).forEach(key => {
        const docs = groups[key];
        if (docs.length > 1) {
          // Sort by timestamp desc or history length (proxy for "most complete")
          docs.sort((a, b) => {
            const timeA = a.timestamp || 0;
            const timeB = b.timestamp || 0;
            return timeB - timeA;
          });
          // Keep the first one, delete the rest
          for (let i = 1; i < docs.length; i++) {
            batch.delete(serviceUpdateService.collection().doc(docs[i].id));
            removedCount++;
          }
        }
      });

      if (removedCount > 0) {
        await batch.commit();
        alert(`Cleanup complete! Removed ${removedCount} duplicate record(s).`);
      } else {
        alert('No duplicates found.');
      }
    } catch (err) {
      console.error('[Cleanup] Error:', err);
      alert('Cleanup failed: ' + err.message);
    } finally {
      showLoading(false);
    }
  };

  if (elements.addServiceUpdateBtn) {
    elements.addServiceUpdateBtn.addEventListener('click', () => {
      clearServiceUpdateForm();
      elements.serviceUpdateModal.show();
    });
  }

  // Service Update filters and chart view listeners
  if (elements.serviceStartDateFilter) {
    elements.serviceStartDateFilter.addEventListener('change', () => { ServiceUpdatesFeatureInstance.setPage(1); try { renderSurChart(); } catch (_) { } });
  }
  if (elements.serviceEndDateFilter) {
    elements.serviceEndDateFilter.addEventListener('change', () => { ServiceUpdatesFeatureInstance.setPage(1); try { renderSurChart(); } catch (_) { } });
  }
  if (elements.serviceProviderFilter) {
    elements.serviceProviderFilter.addEventListener('change', () => { ServiceUpdatesFeatureInstance.setPage(1); try { renderSurChart(); } catch (_) { } });
  }

  if (elements.serviceYearFilter) {
    const currentYear = new Date().getFullYear();
    // Generate years from 2011 to current year (2026 as per user request)
    (function populateServiceUpdateYears() {
      const startYear = 2011;
      const endYear = 2026; 
      const currentSelection = elements.serviceYearFilter.value || currentYear.toString();
      let optionsHtml = '';
      for (let y = endYear; y >= startYear; y--) {
        optionsHtml += `<option value="${y}" ${y.toString() === currentSelection ? 'selected' : ''}>${y}</option>`;
      }
      elements.serviceYearFilter.innerHTML = optionsHtml;
    })();

    elements.serviceYearFilter.addEventListener('change', () => {
      ServiceUpdatesFeatureInstance.setPage(1);
      subscribeServiceUpdates(); // Re-subscribe with new year filter
      try { renderSurChart(); } catch (_) { }
    });
  }

  if (elements.surViewChartToggle) {
    elements.surViewChartToggle.addEventListener('change', () => {
      const isChart = elements.surViewChartToggle.checked;
      if (elements.surTableView) elements.surTableView.style.display = isChart ? 'none' : 'block';
      if (elements.surChartView) elements.surChartView.style.display = isChart ? 'block' : 'none';
      if (isChart) try { renderSurChart(); } catch (_) { }
    });
  }

  signupForm.addEventListener('submit', async e => {
    e.preventDefault();
    const fullName = document.getElementById('signupFullName').value.trim();
    const email = document.getElementById('signupEmail').value.trim();
    const pw = document.getElementById('signupPassword').value;
    const designation = document.getElementById('signupDesignation').value.trim();
    const department = document.getElementById('signupDepartment').value;
    const deptOtherVal = document.getElementById('signupDepartmentOther')?.value.trim();

    if (department === 'Other' && !deptOtherVal) { alert('Please specify your department'); return; }
    const finalDept = department === 'Other' ? deptOtherVal : department;

    try {
      const cred = await auth.createUserWithEmailAndPassword(email, pw);
      await userService.setUserDoc(cred.user.uid, {
        email: email,
        fullName,
        designation,
        department: finalDept,
        approved: false,
        createdAt: serverTimestamp()
      });

      // Queue "Registration Pending" email
      emailService.collection().add({
        to: [email],
        message: {
          subject: 'Registration Pending Approval - Construction Dashboard',
          html: `<p>Dear ${fullName},</p>
                 <p>Thank you for registering. Your account is currently <strong>pending for approval</strong> by the system administrator.</p>
                 <p>You will receive another email once your registration has been approved.</p>
                 <p><em>Reminder: Accessing the system requires a connection to <strong>MWSS_NET</strong>.</em></p>
                 <p>Best regards,<br>System Administrator</p>`
        }
      }).catch(err => console.error('[Signup] Email queue error:', err));

      // Notify Administrator
      emailService.collection().add({
        to: ['johnlowel.fradejas@mwss.gov.ph'],
        message: {
          subject: 'New User Registration - Action Required',
          html: `<h3>New Registration Request</h3>
                 <p>A new user has registered and is pending approval:</p>
                 <ul>
                   <li><strong>Name:</strong> ${fullName}</li>
                   <li><strong>Email:</strong> ${email}</li>
                   <li><strong>Designation:</strong> ${designation}</li>
                   <li><strong>Department:</strong> ${finalDept}</li>
                 </ul>
                 <p>Please log in to the <a href="http://172.16.9.15/construction/">Construction Dashboard</a> to review and approve this account.</p>`
        }
      }).catch(err => console.error('[Signup-Admin] Email queue error:', err));

      alert('Account request submitted. An admin will review your request.');
      auth.signOut();
      showLoginForm();
    } catch (err) {
      alert(err.message);
    }
  });


  function showLogin() {
    loginScreen.classList.remove('d-none');
    loginScreen.style.display = 'flex';
    appWrapper.style.display = 'none';
    logoutBtn.style.display = 'none';
  }

  function showSignup() {
    loginForm.classList.add('d-none');
    loginLinks.classList.add('d-none');
    signupForm.classList.remove('d-none');
    signupLinks.classList.remove('d-none');
    document.getElementById('loginTitle').textContent = 'Sign Up';
  }

  function showLoginForm() {
    signupForm.classList.add('d-none');
    signupLinks.classList.add('d-none');
    loginForm.classList.remove('d-none');
    loginLinks.classList.remove('d-none');
    document.getElementById('loginTitle').textContent = 'Login';
  }

  // link listeners
  if (showSignupLink) showSignupLink.addEventListener('click', e => { e.preventDefault(); showSignup(); });
  if (showLoginLink) showLoginLink.addEventListener('click', e => { e.preventDefault(); showLoginForm(); });

  // Forgot password
  const forgotPwLink = document.getElementById('forgotPwLink');
  if (forgotPwLink) {
    forgotPwLink.addEventListener('click', async e => {
      e.preventDefault();
      // Try to pre-fill with whatever is typed in the email field, if available
      const loginEmailInput = document.querySelector('#loginForm input[type="email"]');
      const defaultEmail = loginEmailInput ? loginEmailInput.value.trim() : '';
      const email = prompt('Enter your registered email address:', defaultEmail);
      if (!email) return;
      try {
        await auth.sendPasswordResetEmail(email.trim());
        alert('Password-reset email sent. Please check your inbox (and spam folder).');
      } catch (err) {
        alert('Failed to send reset email: ' + err.message);
      }
    });
  }

  function showApp() {
    if (isViewOnly) {
      loginScreen.classList.add('d-none');
      appWrapper.style.display = 'block';
      loginBtn.style.display = 'inline-block';
      logoutBtn.style.display = 'none';
      return;
    }
    if (isViewOnly) {
      loginScreen.classList.add('d-none');
      appWrapper.style.display = 'block';
      logoutBtn.style.display = 'none';
      return;
    }
    loginScreen.classList.add('d-none');
    appWrapper.style.display = 'block';
    loginBtn.style.display = 'none';
    logoutBtn.style.display = 'inline-block';
  }


  // Monitor auth state
  auth.onAuthStateChanged(async user => {
    const seq = ++authStateSeq;
    const isCurrentUser = () => (
      seq === authStateSeq &&
      auth.currentUser &&
      auth.currentUser.uid === user.uid &&
      !auth.currentUser.isAnonymous
    );

    if (user) {
      if (user.isAnonymous) {
        isViewOnly = true;
        isAdmin = false;
        isLevel2 = false;
        elevatedAccess = false;

        showApp();
        updateAdminUI();
        updateViewOnlyUI();
        // Default landing: Dashboard
        try { showDashboardSection(); } catch (_) { /* noop */ }
        subscribeDeepwells();
        subscribeServiceUpdates();
        if (unsubscribeProjects) unsubscribeProjects();
        unsubscribeProjects = projectService.subscribe(snap => {
          ProjectsFeatureInstance.setProjects(snap.docs.map(d => ({ id: d.id, ...d.data() })));
          renderProjects();
          try { renderDashboard(); } catch (_) { /* noop */ }
        });
        return;
      }
      isViewOnly = false;
      const emailLower = (user.email || '').trim().toLowerCase();
      isAdmin = emailLower === ADMIN_EMAIL_LOWER;
      console.log('Signed-in as', user.email, '| isAdmin?', isAdmin);
      // Determine user access level
      if (isAdmin) {
        isLevel2 = false;
        elevatedAccess = true;
      } else {
        // Fetch the user's approval doc by UID
        let doc;
        let docRef;
        try {
          docRef = userService.docRef ? userService.docRef(user.uid) : db.collection('users').doc(user.uid);
          doc = await userService.getById(user.uid);
        } catch (err) {
          if (!isCurrentUser()) return;
          console.error('Failed to load user profile', err);
          alert('Unable to verify your account profile right now. Please try again in a moment.');
          try { await auth.signOut(); } catch (_) { /* ignore */ }
          return;
        }
        if (!isCurrentUser()) return;
        // If missing or not approved, attempt a one-time migration by matching on email
        if (!doc.exists || doc.data().approved !== true) {
          try {
            const emailExact = (user.email || '').trim();
            const emailLower = emailExact.toLowerCase();
            let found = await userService.findByEmail(emailExact);
            if (!isCurrentUser()) return;
            if (found.empty) {
              found = await userService.findByEmail(emailLower);
              if (!isCurrentUser()) return;
            }
            if (!found.empty) {
              const src = found.docs[0];
              const data = src.data() || {};
              if (data.approved === true) {
                // Migrate data to UID document
                await userService.setUserDoc(user.uid, {
                  email: emailExact,
                  fullName: data.fullName || '',
                  designation: data.designation || '',
                  department: data.department || '',
                  accessLevel: data.accessLevel || 1,
                  approved: true,
                  updatedAt: serverTimestamp(),
                }, { merge: true });
                // Optional cleanup: delete the non-UID document to avoid duplicates
                if (src.id !== user.uid) {
                  try { await userService.deleteUser(src.id); } catch (_) { /* ignore */ }
                }
              }
            }
          } catch (_) { /* fall through to pending check */ }
          // Re-read after migration attempt
          doc = await docRef.get();
          if (!isCurrentUser()) return;
        }
        if (!doc.exists || doc.data().approved !== true) {
          alert('Your account is pending approval.');
          auth.signOut();
          return;
        }
        const lvl = doc.data().accessLevel || 1;
        isLevel2 = lvl === 2;
        elevatedAccess = isLevel2; // admins already handled above
      }
      if (!isCurrentUser()) return;
      showApp();
      updateAdminUI();
      updateViewOnlyUI();
      // Default landing: Dashboard
      try { showDashboardSection(); } catch (_) { /* noop */ }
      // Ensure approved users list is available for chat recipients even for non-admins
      if (!isAdmin) {
        loadApprovedUsers();
      }
      subscribeDeepwells();
      subscribeServiceUpdates();
      subscribePersonalTasks();
      subscribePresentations();
      subscribeOpcr();
      subscribeIpcr();
      // Start notification listener for smart reminders
      if (notificationsFeature) notificationsFeature.startListening(user);
      if (unsubscribeProjects) unsubscribeProjects();
      const legacyKey = 'construction_projects';
      async function migrateLegacyIfAny() {
        try {
          const legacy = JSON.parse(localStorage.getItem(legacyKey) || '[]');
          if (!legacy.length) return false;
          const batch = projectService.batch();
          legacy.forEach(p => {
            const id = p.id || (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36));
            batch.set(projectService.collection().doc(id), { ...p, id });
          });
          await batch.commit();
          localStorage.removeItem(legacyKey);
          console.log('Migrated', legacy.length, 'legacy projects to Firestore');
          return true;
        } catch (err) { console.error('Legacy migration failed', err); return false; }
      }

      await migrateLegacyIfAny();
      if (!isCurrentUser()) return;

      unsubscribeProjects = projectService.subscribe(async snap => {
        ProjectsFeatureInstance.setProjects(snap.docs.map(d => ({ id: d.id, ...d.data() })));
        if (ProjectsFeatureInstance.getProjects().length === 0) {
          // attempt legacy localStorage migration
          try {
            const legacy = JSON.parse(localStorage.getItem('construction_projects') || '[]');
            if (legacy.length) {
              const batch = projectService.batch();
              legacy.forEach(p => {
                const id = p.id || (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36));
                batch.set(projectService.collection().doc(id), { ...p, id });
              });
              await batch.commit();
              localStorage.removeItem('construction_projects');
              console.log('Migrated', legacy.length, 'projects from localStorage to Firestore');
            }
          } catch (err) { console.warn('Legacy migration failed', err); }
        }
        renderProjects();
        try { renderDashboard(); } catch (_) { /* noop */ }

        // One-time auto-prompt for cleanup (added as per user request)
        if (isAdmin && !localStorage.getItem('sur_cleanup_auto_done')) {
          setTimeout(async () => {
            if (typeof window.cleanUpDuplicateServiceUpdates === 'function') {
              await window.cleanUpDuplicateServiceUpdates();
              localStorage.setItem('sur_cleanup_auto_done', 'true');
            }
          }, 3000);
        }
      });
    } else {
      // No user signed in: stay on the login screen until the user authenticates.
      isAdmin = false;
      isViewOnly = false;
      isLevel2 = false;
      elevatedAccess = false;
      showLogin();
      updateAdminUI();
      // Clean up any active listeners from previous sessions
      if (unsubscribeProjects) { unsubscribeProjects(); unsubscribeProjects = null; }
      if (unsubscribeDeepwells) { unsubscribeDeepwells(); unsubscribeDeepwells = null; }
      if (unsubscribeServiceUpdates) { unsubscribeServiceUpdates(); unsubscribeServiceUpdates = null; }
      if (unsubscribePersonalTasks) { unsubscribePersonalTasks(); unsubscribePersonalTasks = null; }
      if (unsubscribePresentations) { unsubscribePresentations(); unsubscribePresentations = null; }
      if (unsubscribeOpcr) { unsubscribeOpcr(); unsubscribeOpcr = null; }
      if (unsubscribeIpcr) { unsubscribeIpcr(); unsubscribeIpcr = null; }
      // Stop notification listener
      if (notificationsFeature) notificationsFeature.stopListening();
      ProjectsFeatureInstance.setProjects([]);
      deepwells = [];
      DeepwellsFeatureInstance.setDeepwells([]);
      serviceUpdates = [];
      ServiceUpdatesFeatureInstance.setServiceUpdates([]);
      presentations = [];
      PresentationsFeatureInstance.setPresentations([]);
      personalTasks = [];
      PersonalTasksFeatureInstance.setPersonalTasks([]);
      opcrEntries = [];
      OpcrFeatureInstance.setOpcrEntries([]);
      ipcrEntries = [];
      IpcrFeatureInstance.setIpcrEntries([]);
      // Subscriptions
      subscribeProjects();
      subscribeDeepwells();
      subscribeServiceUpdates();
      subscribePresentations();
      subscribePersonalTasks();
      subscribeOpcr();
      subscribeIpcr();

      try { renderPresentations(); } catch (_) {/* noop */ }
      try { renderOpcr(); } catch (_) {/* noop */ }
      try { renderIpcr(); } catch (_) {/* noop */ }
      renderProjects();
      renderDeepwells();
      renderDeepwellMonthlyChart();
      try { renderPersonalTasks(); } catch (_) { }
      if (typeof renderServiceUpdates === 'function') { try { renderServiceUpdates(); } catch (_) { /* noop */ } }
      try { renderDashboard(); } catch (_) { /* noop */ }
    }
  });

  // Handle login form submission
  loginForm.addEventListener('submit', e => {
    e.preventDefault();
    const email = document.getElementById('loginEmail').value.trim();
    const pw = document.getElementById('loginPassword').value;
    auth.signInWithEmailAndPassword(email, pw)
      .catch(err => alert(err.message));
  });

  // Handle logout
  logoutBtn.addEventListener('click', () => auth.signOut());
  if (loginBtn) loginBtn.addEventListener('click', () => {
    // exit view-only: go back to login page
    isViewOnly = false;
    auth.signOut().catch(() => { }); // ensure no anon session remains
    showLogin();
  });
  // Admin manage pending users
  elements.manageUsersBtn.addEventListener('click', () => {
    if (!isAdmin) return;
    loadPendingUsers();
    loadApprovedUsers();
    bootstrap.Modal.getOrCreateInstance(document.getElementById('usersModal')).show();
  });

  // ==================== Apps Dropdown Menu ====================
  (function initAppsDropdown() {
    const appsToggle = document.getElementById('appsToggle');
    const appsMenu = document.getElementById('appsMenu');
    const headerPersonalTasksBtn = document.getElementById('headerPersonalTasksBtn');
    const headerManageUsersBtn = document.getElementById('headerManageUsersBtn');
    const appsPendingBadge = document.getElementById('appsPendingBadge');

    if (!appsToggle || !appsMenu) return;

    // Toggle dropdown on click
    appsToggle.addEventListener('click', (e) => {
      e.stopPropagation();
      const isOpen = appsMenu.style.display === 'block';
      appsMenu.style.display = isOpen ? 'none' : 'block';
      appsToggle.setAttribute('aria-expanded', !isOpen);
    });

    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
      if (!appsMenu.contains(e.target) && e.target !== appsToggle) {
        appsMenu.style.display = 'none';
        appsToggle.setAttribute('aria-expanded', 'false');
      }
    });

    // Connect Personal Tasks button to existing functionality
    if (headerPersonalTasksBtn) {
      headerPersonalTasksBtn.addEventListener('click', () => {
        appsMenu.style.display = 'none'; // Close dropdown
        if (typeof showPersonalTasksSection === 'function') {
          showPersonalTasksSection();
        } else if (elements.personalTasksBtn) {
          elements.personalTasksBtn.click();
        }
      });
    }

    // Connect Manage Users button to existing functionality
    if (headerManageUsersBtn) {
      headerManageUsersBtn.addEventListener('click', () => {
        appsMenu.style.display = 'none'; // Close dropdown
        if (!isAdmin) return;
        loadPendingUsers();
        loadApprovedUsers();
        bootstrap.Modal.getOrCreateInstance(document.getElementById('usersModal')).show();
      });
    }

    // Sync header Manage Users visibility with admin status
    function syncHeaderManageUsers() {
      if (headerManageUsersBtn) {
        headerManageUsersBtn.style.display = isAdmin ? 'flex' : 'none';
      }
      // Sync pending badge
      if (appsPendingBadge && elements.pendingBadge) {
        const count = elements.pendingBadge.textContent || '0';
        const showBadge = elements.pendingBadge.style.display !== 'none';
        appsPendingBadge.textContent = count;
        appsPendingBadge.style.display = showBadge ? 'inline-flex' : 'none';
      }
    }

    // Initial sync and re-sync on auth state change
    auth.onAuthStateChanged(() => setTimeout(syncHeaderManageUsers, 500));
    syncHeaderManageUsers();
  })();

  // ==================== User Profile Dropdown ====================
  (function initUserProfileDropdown() {
    const userProfileToggle = document.getElementById('userProfileToggle');
    const userProfileMenu = document.getElementById('userProfileMenu');
    const headerUserInitials = document.getElementById('headerUserInitials');
    const menuUserInitials = document.getElementById('menuUserInitials');
    const menuUserName = document.getElementById('menuUserName');
    const menuUserEmail = document.getElementById('menuUserEmail');
    const headerLogoutBtn = document.getElementById('headerLogoutBtn');
    let profileUpdateSeq = 0;

    if (!userProfileToggle || !userProfileMenu) return;

    // Generate initials from full name (e.g., "John Lowel Fradejas" -> "JL")
    function getInitials(fullName) {
      if (!fullName) return 'U';
      const names = fullName.trim().split(/\s+/);
      if (names.length >= 2) {
        return (names[0][0] + names[1][0]).toUpperCase();
      } else if (names.length === 1 && names[0].length > 0) {
        return names[0].substring(0, 2).toUpperCase();
      }
      return 'U';
    }

    function setViewerProfile() {
      const logoutBtnIcon = headerLogoutBtn ? headerLogoutBtn.querySelector('i') : null;
      const logoutBtnText = headerLogoutBtn ? headerLogoutBtn.querySelector('span') : null;

      if (headerUserInitials) headerUserInitials.textContent = 'G';
      if (menuUserInitials) menuUserInitials.textContent = 'G';
      if (menuUserName) menuUserName.textContent = 'Guest Viewer';
      if (menuUserEmail) menuUserEmail.textContent = 'Logged in as viewer';
      if (logoutBtnIcon) logoutBtnIcon.className = 'fa-solid fa-right-to-bracket';
      if (logoutBtnText) logoutBtnText.textContent = 'Back to Login Page';
    }
    window.__setViewerProfile = setViewerProfile;

    // Update user profile display
    function updateUserProfile(user) {
      const updateSeq = ++profileUpdateSeq;
      if (!user) {
        if (isViewOnly) setViewerProfile();
        return;
      }

      const canApplyUserProfile = () => (
        updateSeq === profileUpdateSeq &&
        !isViewOnly &&
        auth.currentUser &&
        auth.currentUser.uid === user.uid &&
        !auth.currentUser.isAnonymous
      );

      const logoutBtnIcon = headerLogoutBtn ? headerLogoutBtn.querySelector('i') : null;
      const logoutBtnText = headerLogoutBtn ? headerLogoutBtn.querySelector('span') : null;

      // Handle anonymous / viewer mode
      if (isViewOnly || user.isAnonymous) {
        setViewerProfile();
        return;
      }

      // Authenticated user: restore logout button label
      if (logoutBtnIcon) logoutBtnIcon.className = 'fa-solid fa-arrow-right-from-bracket';
      if (logoutBtnText) logoutBtnText.textContent = 'Logout';


      // Get user data from Firestore or use email
      const email = user.email || '';
      let fullName = '';

      // Try to get full name from user service
      if (typeof userService !== 'undefined' && userService.findByEmail) {
        userService.findByEmail(email).then(snap => {
          if (!canApplyUserProfile()) return;
          if (!snap.empty) {
            const userData = snap.docs[0].data();
            fullName = userData.fullName || '';
            const initials = getInitials(fullName);

            if (headerUserInitials) headerUserInitials.textContent = initials;
            if (menuUserInitials) menuUserInitials.textContent = initials;
            if (menuUserName) menuUserName.textContent = fullName || email.split('@')[0];
            if (menuUserEmail) menuUserEmail.textContent = email;
          } else {
            // Fallback to email
            const initials = getInitials(email.split('@')[0]);
            if (headerUserInitials) headerUserInitials.textContent = initials;
            if (menuUserInitials) menuUserInitials.textContent = initials;
            if (menuUserName) menuUserName.textContent = email.split('@')[0];
            if (menuUserEmail) menuUserEmail.textContent = email;
          }
        }).catch(() => {
          if (!canApplyUserProfile()) return;
          // Fallback on error
          const initials = getInitials(email.split('@')[0]);
          if (headerUserInitials) headerUserInitials.textContent = initials;
          if (menuUserInitials) menuUserInitials.textContent = initials;
          if (menuUserName) menuUserName.textContent = email.split('@')[0];
          if (menuUserEmail) menuUserEmail.textContent = email;
        });
      } else {
        if (!canApplyUserProfile()) return;
        // No user service, use email
        const initials = getInitials(email.split('@')[0]);
        if (headerUserInitials) headerUserInitials.textContent = initials;
        if (menuUserInitials) menuUserInitials.textContent = initials;
        if (menuUserName) menuUserName.textContent = email.split('@')[0];
        if (menuUserEmail) menuUserEmail.textContent = email;
      }
    }

    // Toggle dropdown on click
    userProfileToggle.addEventListener('click', (e) => {
      e.stopPropagation();
      const isOpen = userProfileMenu.style.display === 'block';
      userProfileMenu.style.display = isOpen ? 'none' : 'block';
      userProfileToggle.setAttribute('aria-expanded', !isOpen);
    });

    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
      if (!userProfileMenu.contains(e.target) && e.target !== userProfileToggle && !userProfileToggle.contains(e.target)) {
        userProfileMenu.style.display = 'none';
        userProfileToggle.setAttribute('aria-expanded', 'false');
      }
    });

    // Logout handler
    if (headerLogoutBtn) {
      headerLogoutBtn.addEventListener('click', () => {
        userProfileMenu.style.display = 'none';
        if (isViewOnly || auth.currentUser?.isAnonymous) {
          isViewOnly = false;
          auth.signOut().catch(() => {});
          showLogin();
        } else {
          auth.signOut();
        }

      });
    }

    // Update profile on auth state change
    auth.onAuthStateChanged((user) => {
      if (user || isViewOnly) {
        setTimeout(() => updateUserProfile(user), 300);
      }
    });
  })();

  // ==================== Global Header Title Manager ====================
  (function initGlobalHeaderTitle() {
    const titleEl = document.getElementById('globalModuleTitle');
    if (!titleEl) return;

    // Map of section IDs to titles
    const sectionTitles = {
      dashboardSection: 'Project Monitoring Dashboard',
      projectsSection: 'Construction Projects',
      reforestationSection: 'Reforestation Projects',
      deepwellsSection: 'Deepwell Monitoring',
      personalTasksSection: 'Personal Task Monitoring',
      serviceUpdateSection: 'Service Update Report',
      presentationsSection: 'Product Presentations',
      calendarSection: 'Calendar of Activities',
      opcrSection: 'OPCR Dashboard',
      opcrFormSection: 'OPCR Dashboard',
      ipcrSection: 'IPCR Dashboard',
      documentRegistrySection: 'Document Registry'
    };

    // Update title based on visible section
    function updateGlobalTitle() {
      for (const [sectionId, title] of Object.entries(sectionTitles)) {
        const section = document.getElementById(sectionId);
        if (section && section.style.display !== 'none') {
          titleEl.textContent = title;
          return;
        }
      }
    }

    // Observe section visibility changes
    const observer = new MutationObserver(() => {
      setTimeout(updateGlobalTitle, 50);
    });

    Object.keys(sectionTitles).forEach(sectionId => {
      const section = document.getElementById(sectionId);
      if (section) {
        observer.observe(section, { attributes: true, attributeFilter: ['style'] });
      }
    });

    // Initial update
    updateGlobalTitle();
  })();

  // Handle window resize/zoom changes - debounced chart refresh
  let resizeTimeout;
  window.addEventListener('resize', () => {
    clearTimeout(resizeTimeout);
    resizeTimeout = setTimeout(() => {
      // Re-render all charts that may be affected by zoom
      try { if (typeof renderSurChartDash === 'function') renderSurChartDash(); } catch (_) { }
      try { if (typeof renderDeepwellMonthlyChart === 'function') renderDeepwellMonthlyChart(); } catch (_) { }
      try { if (typeof renderDeepwellMapDash === 'function') renderDeepwellMapDash(); } catch (_) { }
      try { if (typeof renderDamChart === 'function') renderDamChart(); } catch (_) { }
    }, 250);
  });


  // Map Modal events
  const mapModalEl = document.getElementById('deepwellMapModal');
  if (mapModalEl) {
    mapModalEl.addEventListener('shown.bs.modal', () => {
      renderDeepwellMapDetail();
    });
    mapModalEl.addEventListener('hidden.bs.modal', () => {
      // Destroy detail map when modal closes to free resources
      dashDeepwellMapDetail = destroyMap(dashDeepwellMapDetail, elements.deepwellMapDetail);
    });
  }

  // Keyboard accessibility for map card
  elements.deepwellMapCard?.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      elements.deepwellMapCard.click();
    }
  });

})();
