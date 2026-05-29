// Deepwells feature module: handles CRUD, rendering, pagination, chart, and view/edit modals.
// Depends on AppUtils (fmtNum), DeepwellService, Firebase auth, and DOM elements passed in.
(function (window) {
  if (window.DeepwellsFeature) return;

  function init(opts) {
    const {
      elements,
      deepwellService,
      auth,
      utils,
      perPage = (window.AppConstants && window.AppConstants.DEEPWELLS_PER_PAGE) || 30,
      elevatedAccessRef = () => false,
      isAdminRef = () => false,
    } = opts || {};
    if (!elements || !deepwellService || !auth || !utils) {
      throw new Error('DeepwellsFeature.init missing required dependencies');
    }
    const { fmtNum } = utils;
    let deepwells = [];
    let deepwellsPage = 1;
    let dwMonthlyChart = null;

    function getDmsDigits(raw, maxDigits) {
      return String(raw || '').replace(/\D/g, '').slice(0, maxDigits);
    }

    function getDmsDegreeDigits(digits, maxDegreeDigits, maxDegree) {
      if (digits.length === maxDegreeDigits + 5) {
        const maxDigitDegree = parseInt(digits.slice(0, maxDegreeDigits), 10);
        if (maxDigitDegree > maxDegree) return maxDegreeDigits - 1;
      }
      return maxDegreeDigits;
    }

    function formatDmsValue(raw, maxDegreeDigits, maxDegree) {
      const maxDigits = maxDegreeDigits + 6;
      const digits = getDmsDigits(raw, maxDigits);
      if (!digits) return '';

      const degreeDigits = getDmsDegreeDigits(digits, maxDegreeDigits, maxDegree);
      const degree = digits.slice(0, Math.min(degreeDigits, digits.length));
      if (digits.length <= degreeDigits) return `${degree}°`;

      const minute = digits.slice(degreeDigits, Math.min(degreeDigits + 2, digits.length));
      if (digits.length <= degreeDigits + 2) return `${degree}° ${minute}'`;

      const secondRaw = digits.slice(degreeDigits + 2, degreeDigits + 6);
      const secondWhole = secondRaw.slice(0, 2);
      const secondDecimal = secondRaw.slice(2, 4);
      const second = secondDecimal ? `${secondWhole}.${secondDecimal}` : secondWhole;
      return `${degree}° ${minute}' ${second}"`;
    }

    function normalizeDmsValue(raw, maxDegreeDigits, maxDegree) {
      const value = String(raw || '').trim();
      if (!value) return '';

      const parsed = value.match(/^(\d{1,3})°\s*(\d{1,2})'\s*(\d{1,2})(?:\.(\d{0,2}))?"$/);
      if (parsed) {
        const degree = parsed[1];
        const minute = parsed[2].padStart(2, '0');
        const second = parsed[3].padStart(2, '0');
        const decimal = (parsed[4] || '').padEnd(2, '0');
        return `${degree}° ${minute}' ${second}.${decimal}"`;
      }

      const maxDigits = maxDegreeDigits + 6;
      const digits = getDmsDigits(value, maxDigits);
      if (!digits) return '';
      const degreeDigits = getDmsDegreeDigits(digits, maxDegreeDigits, maxDegree);
      if (digits.length < degreeDigits + 4) {
        return formatDmsValue(digits, maxDegreeDigits, maxDegree);
      }
      const fullDigits = digits.padEnd(degreeDigits + 6, '0');
      return formatDmsValue(fullDigits, maxDegreeDigits, maxDegree);
    }

    function validateDmsValue(value, maxDegree, label) {
      if (!value) return '';
      const match = String(value).trim().match(/^(\d{1,3})°\s*(\d{2})'\s*(\d{2})\.(\d{2})"$/);
      if (!match) return `${label} must be in DMS format, for example 14° 35' 24.12".`;

      const degree = parseInt(match[1], 10);
      const minute = parseInt(match[2], 10);
      const second = parseFloat(`${match[3]}.${match[4]}`);
      if (degree < 0 || degree > maxDegree) return `${label} degrees must be from 0 to ${maxDegree}.`;
      if (minute < 0 || minute > 59) return `${label} minutes must be from 00 to 59.`;
      if (second < 0 || second >= 60) return `${label} seconds must be from 00.00 to 59.99.`;
      return '';
    }

    function attachDmsFormatter(inputId, maxDegreeDigits, maxDegree) {
      const input = document.getElementById(inputId);
      if (!input) return;
      input.addEventListener('input', () => {
        input.value = formatDmsValue(input.value, maxDegreeDigits, maxDegree);
        input.setSelectionRange(input.value.length, input.value.length);
      });
      input.addEventListener('blur', () => {
        input.value = normalizeDmsValue(input.value, maxDegreeDigits, maxDegree);
      });
    }

    const isViewOnly = () => {
      if (typeof window.__VIEW_ONLY__ === 'function') return !!window.__VIEW_ONLY__();
      return !!window.__VIEW_ONLY__ || false;
    };

    function setDeepwells(list) {
      deepwells = Array.isArray(list) ? list.slice() : [];
      // Auto-refresh statuses on data load, but only if user has elevated access
      // and not in view-only mode to prevent infinite loops for viewers.
      if (elevatedAccessRef() && !isViewOnly()) {
        refreshAllDeepwellStatuses();
      }
    }

    function getDeepwells() {
      return deepwells.slice();
    }

    /**
     * Generate an array of YYYY-MM keys for the last N months
     * counting back from (and including) the provided startMonth or current month.
     * e.g. if startMonth is 2026-03 and n=2 → ['2026-02','2026-03']
     */
    function generateMonthKeys(n, startMonth) {
      const keys = [];
      let refDate = new Date();
      if (startMonth && typeof startMonth === 'string') {
        const parts = startMonth.trim().split('-');
        if (parts.length >= 2) {
          const y = parseInt(parts[0], 10);
          const m = parseInt(parts[1], 10);
          if (!isNaN(y) && !isNaN(m)) {
            refDate = new Date(y, m - 1, 1);
          }
        }
      }
      for (let i = 0; i < n; i++) {
        const d = new Date(refDate.getFullYear(), refDate.getMonth() - i, 1);
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        keys.push(`${y}-${m}`);
      }
      return keys;
    }

    /** Find the latest month (YYYY-MM) per provider (MWCI vs MWSI) */
    function getLatestMonthsByProvider(extraMonths = []) {
      const latest = { MWCI: '', MWSI: '' };
      const check = (m, prov) => {
        if (!m || !m.month) return;
        const val = String(m.month).trim();
        const p = (prov || '').toUpperCase();
        if (latest[p] !== undefined && val > latest[p]) {
          latest[p] = val;
        }
      };

      deepwells.forEach(dw => {
        const prov = (dw.provider || '').toUpperCase();
        (dw.months || []).forEach(m => check(m, prov));
      });

      // Handle any extra months passed in (e.g. from a form currently being saved)
      extraMonths.forEach(m => {
        // If we don't know the provider for the extra months, we might need to be careful,
        // but typically this is called within save context where we know the provider.
        // For now, if provider isn't passed, check if it's a valid month string.
        if (m && m.month && /^\d{4}-\d{2}/.test(m.month)) {
          // We apply to both if unknown, or better: this function is mostly used on the full list.
          const val = String(m.month).trim();
          if (val > latest.MWCI) latest.MWCI = val;
          if (val > latest.MWSI) latest.MWSI = val;
        }
      });

      const now = new Date();
      const fallback = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
      
      if (!latest.MWCI) latest.MWCI = fallback;
      if (!latest.MWSI) latest.MWSI = fallback;

      return latest;
    }

    /** Find the latest month (YYYY-MM) across all deepwells (Backward compat) */
    function getGlobalLatestMonth(extraMonths = []) {
      const perProv = getLatestMonthsByProvider(extraMonths);
      return perProv.MWCI > perProv.MWSI ? perProv.MWCI : perProv.MWSI;
    }

    /**
     * Determine the auto-computed status for a deepwell.
     * Rules are relative to latestMonth (global latest update):
     *  1. If months has data in ANY of the last 2 months from latestMonth → Active
     *  2. If currently Active and NO data in the last 2 months → Standby
     *  3. If currently Standby and NO data in the last 8 months → Inactive
     * Returns { status, changed } where changed is true if status differs.
     */
    function computeDeepwellAutoStatus(months, currentStatus, latestMonth) {
      const monthSet = new Set((months || []).map(m => String(m.month || '').trim()));
      const last2 = generateMonthKeys(2, latestMonth);
      const hasRecentData = last2.some(k => monthSet.has(k));

      // Use case-insensitive comparison to avoid constant updates due to casing
      const cur = (currentStatus || '').trim();
      const curLower = cur.toLowerCase();

      if (hasRecentData) {
        return { status: 'Active', changed: curLower !== 'active' };
      }

      // No recent data in last 2 months relative to latest update
      if (curLower === 'active' || curLower === '') {
        return { status: 'Standby', changed: true };
      }

      if (curLower === 'standby') {
        const last8 = generateMonthKeys(8, latestMonth);
        const hasAnyIn8 = last8.some(k => monthSet.has(k));
        if (!hasAnyIn8) {
          return { status: 'Inactive', changed: true };
        }
      }

      // If already Inactive or any other status (and no recent data), no auto-change needed
      return { status: cur, changed: false };
    }

    /**
     * On data load, check all deepwells and batch-update any whose
     * status should change per the auto-status rules.
     */
    let refreshTimeout;
    function refreshAllDeepwellStatuses() {
      // Debounce the refresh to prevent rapid-fire updates during data syncing
      clearTimeout(refreshTimeout);
      refreshTimeout = setTimeout(() => {
        // Only Admins (level 1) can commit auto-status changes to avoid conflicts
        if (!isAdminRef()) return;

        const updates = [];
        const latestMonths = getLatestMonthsByProvider();
        deepwells.forEach(dw => {
          const prov = (dw.provider || '').toUpperCase();
          const latestMonth = latestMonths[prov] || latestMonths.MWCI; // Fallback
          const result = computeDeepwellAutoStatus(dw.months, dw.status, latestMonth);
          if (result.changed) {
            const prevStatus = dw.status || '(none)';
            dw.status = result.status;
            const histEntry = {
              email: 'system',
              fullName: 'Auto-Status',
              timestamp: new Date().toISOString(),
              action: `auto-status: ${prevStatus} → ${result.status}`
            };
            if (!Array.isArray(dw.history)) dw.history = [];
            dw.history.push(histEntry);
            updates.push(dw);
          }
        });
        if (updates.length > 0 && deepwellService) {
          // Batch-save changed deepwells to Firebase
          const batch = deepwellService.batch();
          updates.forEach(dw => {
            batch.set(deepwellService.collection().doc(dw.id), { status: dw.status, history: dw.history }, { merge: true });
          });
          batch.commit()
            .then(() => console.log(`[Deepwells] Auto-status updated ${updates.length} deepwell(s)`))
            .catch(err => console.error('[Deepwells] Auto-status batch error', err));
        }
      }, 2000); // 2 second delay to let data stabilize
    }

    function monthKeyToLabel(ym) {
      if (!ym || typeof ym !== 'string' || ym.length < 7) return ym || '';
      const [y, m] = ym.split('-');
      const d = new Date(Number(y), Number(m) - 1, 1);
      if (isNaN(d.getTime())) return ym;
      return d.toLocaleString('en-US', { month: 'short', year: 'numeric' });
    }

    function aggregateDeepwellMonthlyTotals() {
      const byProv = { MWCI: {}, MWSI: {} };
      (deepwells || []).forEach(dw => {
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
        labels: months.map(monthKeyToLabel),
        // Convert cu.m to ML (Million Liters): 1 cu.m = 1000 liters = 0.001 ML
        mwci: months.map(k => +((byProv.MWCI[k] || 0) / 1000).toFixed(2)),
        mwsi: months.map(k => +((byProv.MWSI[k] || 0) / 1000).toFixed(2)),
        hasData: months.length > 0
      };
    }

    function renderDeepwellMonthlyChart() {
      const canvas = elements.deepwellMonthlyChartCanvas;
      const emptyEl = elements.dwChartEmpty;
      if (!canvas) return; // not on this page
      const agg = aggregateDeepwellMonthlyTotals();
      if (!agg.hasData) {
        if (emptyEl) emptyEl.style.display = 'block';
        canvas.style.display = 'none';
        if (dwMonthlyChart) { dwMonthlyChart.destroy(); dwMonthlyChart = null; }
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
          },
          {
            label: 'MWSI',
            data: agg.mwsi,
            borderColor: '#198754',
            backgroundColor: 'rgba(25,135,84,0.15)',
            tension: 0.25,
            fill: true,
          }
        ]
      };
      const isMobile = window.matchMedia('(max-width: 576px)').matches;
      const tickFontSize = isMobile ? 10 : 12;
      const opts = {
        responsive: true,
        maintainAspectRatio: true,
        aspectRatio: isMobile ? 1.6 : 2.6,
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
        scales: {
          x: { title: { display: true, text: 'Month' }, ticks: { font: { size: tickFontSize } }, grid: { color: 'rgba(0,0,0,0.06)' } },
          y: { title: { display: true, text: 'Production (ML)' }, beginAtZero: true, ticks: { font: { size: tickFontSize } }, grid: { color: 'rgba(0,0,0,0.06)' } }
        }
      };
      if (dwMonthlyChart) {
        dwMonthlyChart.data = data;
        dwMonthlyChart.options = opts;
        dwMonthlyChart.update();
      } else {
        const ctx = canvas.getContext('2d');
        dwMonthlyChart = new Chart(ctx, { type: 'line', data, options: opts });
      }
    }

    function deepwellRowHtml(dw) {
      const canEdit = !isViewOnly() && auth.currentUser;
      let actionsHtml = canEdit ? `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="editDeepwell('${dw.id}')"><i class="fa fa-pencil"></i></button>` : '';
      if (isAdminRef()) {
        actionsHtml += `<button class="btn btn-sm btn-danger" title="Delete" onclick="deleteDeepwell('${dw.id}')"><i class="fa fa-trash"></i></button>`;
      }
      return `<tr data-id="${dw.id}">
        <td>${dw.name}</td>
        <td>${dw.provider}</td>
        <td>${dw.permit || ''}</td>
        <td>${dw.status || ''}</td>
        <td>${fmtNum(dw.ratedYield)}</td>
        <td>${fmtNum(dw.avgProd)}</td>
        <td>${fmtNum(dw.totalProd)}</td>
        <td class="d-flex gap-1">${actionsHtml}</td>
      </tr>`;
    }

    function buildDeepwellsPagination(total, page, pageSize) {
      const pager = elements.deepwellsPagination;
      if (!pager) return;
      const totalPages = Math.max(1, Math.ceil(total / pageSize));
      if (totalPages <= 1) { pager.innerHTML = ''; return; }
      const makeBtn = (label, data, { disabled = false, active = false } = {}) => {
        const cls = `page-item${disabled ? ' disabled' : ''}${active ? ' active' : ''}`;
        const aria = active ? ' aria-current="page"' : '';
        const dp = disabled ? '' : ` data-page="${data}"`;
        return `<li class="${cls}"><button type="button" class="page-link"${aria}${dp}>${label}</button></li>`;
      };
      let html = '';
      html += makeBtn('Prev', 'prev', { disabled: page <= 1 });
      const maxShown = 7;
      let start = Math.max(1, page - Math.floor(maxShown / 2));
      let end = Math.min(totalPages, start + maxShown - 1);
      if (end - start + 1 < maxShown) start = Math.max(1, end - maxShown + 1);
      if (start > 1) {
        html += makeBtn('1', 1, { active: page === 1 });
        if (start > 2) html += `<li class="page-item disabled"><span class="page-link">…</span></li>`;
      }
      for (let i = start; i <= end; i++) {
        if (i === 1 || i === totalPages) continue;
        html += makeBtn(String(i), i, { active: page === i });
      }
      if (end < totalPages) {
        if (end < totalPages - 1) html += `<li class="page-item disabled"><span class="page-link">…</span></li>`;
        html += makeBtn(String(totalPages), totalPages, { active: page === totalPages });
      }
      html += makeBtn('Next', 'next', { disabled: page >= totalPages });
      pager.innerHTML = html;
    }

    function attachRowEvents() {
      const tbody = elements.deepwellsTbody;
      if (!tbody) return;
      Array.from(tbody.querySelectorAll('tr')).forEach(row => {
        row.addEventListener('click', e => {
          if (isViewOnly()) return;
          if (e.target.closest('button')) return;
          const id = row.dataset.id;
          editDeepwell(id, false);
        });
      });
    }

    function renderDeepwells() {
      if (elements.addDeepwellBtn) {
        elements.addDeepwellBtn.style.display = (!isViewOnly() && auth.currentUser) ? 'inline-block' : 'none';
      }
      let list = deepwells.slice();
      const provider = elements.dwProviderFilter?.value;
      const status = elements.dwStatusFilter?.value;
      const query = (elements.dwSearchInput?.value || '').trim().toLowerCase();
      if (provider) list = list.filter(dw => dw.provider === provider);
      if (status) list = list.filter(dw => dw.status === status);
      if (query) list = list.filter(dw => {
        const name = (dw.name || '').toLowerCase();
        const prov = (dw.provider || '').toLowerCase();
        const permit = (dw.permit || '').toLowerCase();
        return name.includes(query) || prov.includes(query) || permit.includes(query);
      });
      const sortKey = elements.dwProdSort?.value || '';
      if (sortKey) {
        const toNum = v => Number(v) || 0;
        const byAvgDesc = (a, b) => toNum(b.avgProd) - toNum(a.avgProd);
        const byAvgAsc = (a, b) => toNum(a.avgProd) - toNum(b.avgProd);
        const byTotDesc = (a, b) => toNum(b.totalProd) - toNum(a.totalProd);
        const byTotAsc = (a, b) => toNum(a.totalProd) - toNum(b.totalProd);
        const cmp = sortKey === 'avgDesc' ? byAvgDesc
          : sortKey === 'avgAsc' ? byAvgAsc
            : sortKey === 'totalDesc' ? byTotDesc
              : sortKey === 'totalAsc' ? byTotAsc
                : null;
        if (cmp) list = [...list].sort(cmp);
      }
      const total = list.length;
      const totalPages = Math.max(1, Math.ceil(total / perPage));
      if (deepwellsPage > totalPages) deepwellsPage = totalPages;
      if (deepwellsPage < 1) deepwellsPage = 1;
      const start = (deepwellsPage - 1) * perPage;
      const end = start + perPage;
      const pageItems = list.slice(start, end);
      const tbody = elements.deepwellsTbody;
      if (tbody) {
        tbody.innerHTML = pageItems.map(deepwellRowHtml).join('');
        attachRowEvents();
      }
      buildDeepwellsPagination(total, deepwellsPage, perPage);
    }

    function gatherDwMonths() {
      const rows = Array.from(elements.dwMonthsBody?.querySelectorAll('tr') || []);
      return rows.map(r => ({
        month: r.querySelector('.dw-month')?.value || '',
        prod: parseFloat(r.querySelector('.dw-prod')?.value || '0') || 0
      })).filter(m => m.month && m.prod);
    }

    function updateDwStats() {
      if (!elements.dwMonthsBody) return;
      const months = gatherDwMonths();
      const total = months.reduce((s, m) => s + m.prod, 0);
      const avg = months.length ? (total / months.length) : 0;
      const totalEl = document.getElementById('dwTotalProd');
      const avgEl = document.getElementById('dwAvgProd');
      if (totalEl) totalEl.value = total ? total.toFixed(2) : '';
      if (avgEl) avgEl.value = avg ? avg.toFixed(2) : '';
    }

    function addDwMonthRow(data = {}) {
      if (!elements.dwMonthsBody) return;
      const tr = document.createElement('tr');
      tr.innerHTML = `<td><input type="month" class="form-control form-control-sm dw-month" value="${data.month || ''}"/></td>
                <td><input type="number" step="0.01" class="form-control form-control-sm dw-prod" value="${data.prod || ''}"/></td>
                <td class="text-end"><button type="button" class="btn btn-sm btn-outline-danger dw-remove"><i class="fa fa-minus"></i></button></td>`;
      tr.querySelector('.dw-remove').onclick = () => { tr.remove(); updateDwStats(); };
      elements.dwMonthsBody.appendChild(tr);
      updateDwStats();
    }

    async function saveDeepwell(dw) {
      await deepwellService.save(dw);
    }

    function resetDeepwellForm() {
      if (elements.deepwellForm) elements.deepwellForm.reset();
      const idEl = document.getElementById('deepwellId');
      if (idEl) idEl.value = '';
      if (elements.dwMonthsBody) {
        elements.dwMonthsBody.innerHTML = '';
        addDwMonthRow();
      }
      updateDwStats();
      const historyContainer = document.getElementById('dwHistoryContainer');
      if (historyContainer) historyContainer.innerHTML = '';
      if (elements.deepwellForm) {
        Array.from(elements.deepwellForm.elements).forEach(el => {
          if (el.tagName === 'INPUT' || el.tagName === 'SELECT') {
            el.readOnly = false;
            el.disabled = false;
            if (el.type === 'text' && ['dwRatedYield', 'dwAvgProd', 'dwTotalProd'].includes(el.id)) {
              el.type = 'number';
            }
          }
        });
      }
      const saveBtn = document.getElementById('saveDeepwellBtn');
      if (saveBtn) saveBtn.style.display = 'inline-block';
    }

    function renderHistory(dw) {
      const historyContainer = document.getElementById('dwHistoryContainer');
      if (!historyContainer) return;
      const hist = Array.isArray(dw.history) ? [...dw.history] : [];
      hist.sort((a, b) => {
        const at = a.timestamp || '';
        const bt = b.timestamp || '';
        return (at > bt ? -1 : at < bt ? 1 : 0);
      });
      if (!isAdminRef() || hist.length === 0) {
        historyContainer.innerHTML = '';
      } else {
        const rows = hist.map(h => {
          const when = h.timestamp ? new Date(h.timestamp).toLocaleString() : '';
          const who = h.fullName || h.email || 'Unknown user';
          const act = h.action || 'update';
          return `<div class="d-flex align-items-start gap-2 py-1">
          <i class="fa-regular fa-clock mt-1 text-muted"></i>
          <div>
            <div><strong>${act}</strong> by ${who}</div>
            <div class="text-muted small">${when}</div>
          </div>
        </div>`;
        }).join('');
        historyContainer.innerHTML = `<div class="small text-muted mt-3">Edit History</div><div class="border rounded p-2 bg-light small">${rows}</div>`;
      }
    }

    function editDeepwell(id, edit = true) {
      const dw = deepwells.find(x => x.id === id);
      if (!dw || !elements.deepwellForm) return;
      const setVal = (id, v) => { const el = document.getElementById(id); if (el) el.value = v || ''; };
      setVal('deepwellId', dw.id);
      setVal('dwName', dw.name);
      setVal('dwProvider', dw.provider);
      setVal('dwPermit', dw.permit);
      setVal('dwStatus', dw.status);
      setVal('dwRatedYield', dw.ratedYield);
      setVal('dwAvgProd', dw.avgProd);
      setVal('dwTotalProd', dw.totalProd);
      setVal('dwLocation', dw.location);
      setVal('dwMunicipality', dw.municipality);
      setVal('dwLatitude', dw.latitude);
      setVal('dwLongitude', dw.longitude);
      if (elements.dwMonthsBody) {
        elements.dwMonthsBody.innerHTML = '';
        (dw.months || []).forEach(m => addDwMonthRow(m));
        if ((dw.months || []).length === 0) addDwMonthRow();
      }
      updateDwStats();
      if (elements.deepwellForm) {
        Array.from(elements.deepwellForm.elements).forEach(el => {
          if (el.tagName === 'INPUT' || el.tagName === 'SELECT') {
            el.readOnly = !edit;
            el.disabled = !edit && el.tagName === 'SELECT';
          }
        });
      }
      ['dwRatedYield', 'dwAvgProd', 'dwTotalProd'].forEach(id => {
        const inp = document.getElementById(id);
        if (!inp) return;
        if (edit) {
          if (inp.type !== 'number') inp.type = 'number';
        } else {
          inp.type = 'text';
          inp.value = fmtNum(inp.value);
        }
      });
      if (!edit && elements.dwMonthsBody) {
        elements.dwMonthsBody.querySelectorAll('.dw-prod').forEach(inp => { inp.type = 'text'; inp.value = fmtNum(inp.value); inp.readOnly = true; });
        elements.dwMonthsBody.querySelectorAll('.dw-month').forEach(inp => { inp.readOnly = true; });
      }
      renderHistory(dw);
      const saveBtn = document.getElementById('saveDeepwellBtn');
      if (saveBtn) saveBtn.style.display = edit ? 'inline-block' : 'none';
      elements.deepwellModal?.show();
    }

    function startCreate() {
      resetDeepwellForm();
      elements.deepwellModal?.show();
    }

    async function onSaveDeepwell(e) {
      e.preventDefault();
      if (!elements.deepwellForm) return;
      const fd = new FormData(elements.deepwellForm);
      const id = fd.get('deepwellId') || (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36));
      const user = auth.currentUser;
      const userEmail = user?.email || 'viewer';
      const userFullName = (user?.displayName || '').trim();
      const existingDw = deepwells.find(dw => dw.id === id);
      const history = existingDw?.history ? [...existingDw.history] : [];
      history.push({ email: userEmail, fullName: userFullName, timestamp: new Date().toISOString(), action: existingDw ? 'edit' : 'create' });
      const months = gatherDwMonths();
      let formStatus = (fd.get('dwStatus') || '').trim();
      const latitude = normalizeDmsValue(fd.get('dwLatitude'), 2, 90);
      const longitude = normalizeDmsValue(fd.get('dwLongitude'), 3, 180);
      const latitudeError = validateDmsValue(latitude, 90, 'Latitude');
      if (latitudeError) { alert(latitudeError); return; }
      const longitudeError = validateDmsValue(longitude, 180, 'Longitude');
      if (longitudeError) { alert(longitudeError); return; }
      const latitudeEl = document.getElementById('dwLatitude');
      const longitudeEl = document.getElementById('dwLongitude');
      if (latitudeEl) latitudeEl.value = latitude;
      if (longitudeEl) longitudeEl.value = longitude;

      // Auto-status determination based on latest update for this provider
      const prov = (fd.get('dwProvider') || '').toUpperCase();
      const latestMonths = getLatestMonthsByProvider(months);
      const latestMonth = latestMonths[prov] || latestMonths.MWCI;
      const autoResult = computeDeepwellAutoStatus(months, formStatus, latestMonth);
      if (autoResult.changed) {
        const prevStatus = formStatus || '(none)';
        formStatus = autoResult.status;
        history.push({
          email: 'system',
          fullName: 'Auto-Status',
          timestamp: new Date().toISOString(),
          action: `auto-status: ${prevStatus} → ${formStatus}`
        });
      }

      const deepwell = {
        id,
        name: (fd.get('dwName') || '').trim(),
        provider: fd.get('dwProvider'),
        permit: (fd.get('dwPermit') || '').trim(),
        status: formStatus,
        ratedYield: parseFloat(fd.get('dwRatedYield')) || 0,
        months,
        avgProd: parseFloat(fd.get('dwAvgProd')) || 0,
        totalProd: parseFloat(fd.get('dwTotalProd')) || 0,
        location: (fd.get('dwLocation') || '').trim(),
        municipality: (fd.get('dwMunicipality') || '').trim(),
        latitude,
        longitude,
        history
      };
      if (!deepwell.name) { alert('Deepwell Name is required'); return; }
      saveDeepwell(deepwell)
        .then(() => {
          if (elements.deepwellModal) elements.deepwellModal.hide();
          resetDeepwellForm();
          deepwells = [...deepwells.filter(dwItem => dwItem.id !== deepwell.id), deepwell];
          renderDeepwells();
          renderDeepwellMonthlyChart();
        })
        .catch(err => alert(err.message));
    }

    async function deleteDeepwell(id) {
      if (!elevatedAccessRef()) { alert('Only admin/level2 can delete deepwells'); return; }
      if (!confirm('Delete this deepwell?')) return;
      try {
        await deepwellService.remove(id);
        deepwells = deepwells.filter(dw => dw.id !== id);
        renderDeepwells();
        renderDeepwellMonthlyChart();
      } catch (err) { alert(err.message); }
    }


    function attachFormListeners() {
      if (elements.deepwellForm) elements.deepwellForm.addEventListener('submit', onSaveDeepwell);
      attachDmsFormatter('dwLatitude', 2, 90);
      attachDmsFormatter('dwLongitude', 3, 180);
      if (elements.addDwMonthBtn) elements.addDwMonthBtn.addEventListener('click', () => addDwMonthRow());
      if (elements.addDeepwellBtn) elements.addDeepwellBtn.addEventListener('click', () => startCreate());
      if (elements.dwMonthsBody) elements.dwMonthsBody.addEventListener('input', updateDwStats);
      const deepwellModalEl = document.getElementById('deepwellModal');
      if (deepwellModalEl) deepwellModalEl.addEventListener('hidden.bs.modal', resetDeepwellForm);
      if (elements.generateDeepwellReportBtn) {
        elements.generateDeepwellReportBtn.addEventListener('click', () => {
          try {
            sessionStorage.setItem('deepwellsData', JSON.stringify(deepwells));
            window.open('deepwell_summary.html', '_blank');
          } catch (err) {
            console.error('Error saving deepwells data to sessionStorage', err);
            alert('Could not open Summary Report: ' + err.message);
          }
        });
      }

      const saveBtn = document.getElementById('saveDeepwellBtn');
      if (saveBtn) {
        saveBtn.addEventListener('click', (e) => {
          if (!elements.deepwellForm.checkValidity()) {
            elements.deepwellForm.reportValidity();
            return;
          }
          e.preventDefault();
          onSaveDeepwell(e);
        });
      }
      elements.dwProviderFilter?.addEventListener('change', () => { deepwellsPage = 1; renderDeepwells(); });
      elements.dwStatusFilter?.addEventListener('change', () => { deepwellsPage = 1; renderDeepwells(); });
      elements.dwProdSort?.addEventListener('change', () => { deepwellsPage = 1; renderDeepwells(); });
      elements.dwSearchInput?.addEventListener('input', () => { deepwellsPage = 1; renderDeepwells(); });
      if (elements.deepwellsPagination) {
        elements.deepwellsPagination.addEventListener('click', (e) => {
          const btn = e.target.closest('button.page-link');
          if (!btn) return;
          const li = btn.closest('li');
          if (li && li.classList.contains('disabled')) return;
          const dp = btn.getAttribute('data-page');
          if (!dp) return;
          if (dp === 'prev') {
            if (deepwellsPage > 1) deepwellsPage--;
          } else if (dp === 'next') {
            deepwellsPage++;
          } else {
            const n = parseInt(dp, 10);
            if (!isNaN(n)) deepwellsPage = n;
          }
          renderDeepwells();
          window.scrollTo({ top: 0, behavior: 'smooth' });
        });
      }
    }

    attachFormListeners();
    if (elements.deepwellsTbody) { /* initial render */ renderDeepwells(); renderDeepwellMonthlyChart(); }

    window.editDeepwell = editDeepwell;
    window.deleteDeepwell = deleteDeepwell;
    window.startDeepwellCreate = startCreate;

    return {
      setDeepwells,
      getDeepwells,
      getLatestMonthsByProvider,
      computeDeepwellAutoStatus,
      generateMonthKeys,
      renderDeepwells,
      renderDeepwellMonthlyChart,
      setPage(n) { deepwellsPage = n; },
    };
  }

  window.DeepwellsFeature = { init };
})(window);
