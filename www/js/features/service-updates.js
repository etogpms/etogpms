// Service Updates feature module: manages list state, filters/pagination, render, and delete permissions.
(function (window) {
  if (window.ServiceUpdatesFeature) return;

  function init(opts) {
    const {
      elements,
      serviceUpdateService,
      auth,
      utils,
      perPage = (window.AppConstants && window.AppConstants.SERVICE_UPDATES_PER_PAGE) || 30,
      elevatedAccessRef = () => false,
      isAdminRef = () => false,
    } = opts || {};
    if (!elements || !serviceUpdateService || !auth || !utils) {
      throw new Error('ServiceUpdatesFeature.init missing required dependencies');
    }
    const { safeDateCompare, fmtNum, formatDateUI } = utils;
    let serviceUpdates = [];
    let serviceUpdatesPage = 1;

    function isViewOnly() {
      if (typeof window.__VIEW_ONLY__ === 'function') return !!window.__VIEW_ONLY__();
      return !!window.__VIEW_ONLY__ || false;
    }

    function setServiceUpdates(list) {
      serviceUpdates = Array.isArray(list) ? list.slice() : [];
    }
    function getServiceUpdates() {
      return serviceUpdates.slice();
    }

    function serviceUpdateRowHtml(s) {
      let actionsHtml = elevatedAccessRef() ? `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="editServiceUpdate('${s.id}', true)"><i class="fa fa-pencil"></i></button>` : '';
      if (isAdminRef()) {
        actionsHtml += `<button class="btn btn-sm btn-danger" title="Delete" onclick="deleteServiceUpdate('${s.id}')"><i class="fa fa-trash"></i></button>`;
      }
      return `<tr data-id="${s.id}">
      <td>${formatDateUI(s.date) || ''}</td>
      <td>${s.provider || ''}</td>
      <td>${(Array.isArray(s.plants) ? s.plants.length : 0)}</td>
      <td>${fmtNum(s.inflows || 0)}</td>
      <td>${fmtNum(s.production || 0)}</td>
      <td>${fmtNum(s.supplyAug || 0)}</td>
      <td>${fmtNum(s.angat || 0)}</td>
      <td>${fmtNum(s.ipo || 0)}</td>
      <td>${fmtNum(s.laMesa || 0)}</td>
      <td>${s.damAsOf || ''}</td>
      <td class="d-flex gap-1">${actionsHtml}</td>
    </tr>`;
    }

    function renderServiceUpdates() {
      if (elements.addServiceUpdateBtn) {
        elements.addServiceUpdateBtn.style.display = (elevatedAccessRef() && auth.currentUser && !isViewOnly()) ? 'inline-block' : 'none';
      }
      const bulkImpBtn = document.getElementById('surBulkImportBtn');
      if (bulkImpBtn) bulkImpBtn.style.display = (elevatedAccessRef() && auth.currentUser && !isViewOnly()) ? 'inline-block' : 'none';
      const bulkTplBtn = document.getElementById('surBulkTemplateBtn');
      if (bulkTplBtn) bulkTplBtn.style.display = 'inline-block';
      const exportBtn = document.getElementById('surExportBtn');
      if (exportBtn) exportBtn.style.display = 'inline-block';

      let list = serviceUpdates.slice();
      const start = elements.serviceStartDateFilter?.value || '';
      const end = elements.serviceEndDateFilter?.value || '';
      const prov = elements.serviceProviderFilter?.value || '';
      if (start) list = list.filter(s => (s.date || '') >= start);
      if (end) list = list.filter(s => (s.date || '') <= end);
      if (prov) list = list.filter(s => (s.provider || '') === prov);
      list.sort((a, b) => {
        const ad = a.date || ''; const bd = b.date || '';
        const cmp = safeDateCompare(ad, bd);
        if (cmp !== 0) return cmp > 0 ? -1 : 1;
        return (a.provider || '').localeCompare(b.provider || '');
      });
      const total = list.length;
      const maxPage = Math.max(1, Math.ceil(total / perPage));
      if (serviceUpdatesPage > maxPage) serviceUpdatesPage = maxPage;
      const startIdx = (serviceUpdatesPage - 1) * perPage;
      const endIdx = Math.min(startIdx + perPage, total);
      const pageSlice = list.slice(startIdx, endIdx);
      const tbody = elements.serviceUpdatesTbody;
      if (!tbody) return;
      tbody.innerHTML = pageSlice.map(serviceUpdateRowHtml).join('');
      Array.from(tbody.querySelectorAll('tr')).forEach(tr => {
        tr.style.cursor = 'pointer';
        tr.addEventListener('click', e => {
          if (e.target.closest('button')) return;
          if (typeof window.viewServiceUpdateDetails === 'function') {
            window.viewServiceUpdateDetails(tr.dataset.id);
          }
        });
      });
      const infoEl = elements.serviceUpdatesPageInfo;
      if (infoEl) {
        if (total === 0) { infoEl.textContent = '0–0 of 0'; }
        else { infoEl.textContent = `${(startIdx + 1).toLocaleString()}–${endIdx.toLocaleString()} of ${total.toLocaleString()}`; }
      }
      if (elements.serviceUpdatesPrev) { elements.serviceUpdatesPrev.disabled = (serviceUpdatesPage <= 1); }
      if (elements.serviceUpdatesNext) { elements.serviceUpdatesNext.disabled = (serviceUpdatesPage >= maxPage); }
    }

    async function deleteServiceUpdate(id) {
      if (!isAdminRef()) { alert('Only admin can delete'); return; }
      if (!confirm('Delete this report?')) return;
      try { await serviceUpdateService.remove(id); } catch (err) { alert(err.message); }
    }

    function attachListeners() {
      if (elements.serviceUpdatesPrev) {
        elements.serviceUpdatesPrev.addEventListener('click', () => { if (serviceUpdatesPage > 1) { serviceUpdatesPage--; renderServiceUpdates(); } });
      }
      if (elements.serviceUpdatesNext) {
        elements.serviceUpdatesNext.addEventListener('click', () => { serviceUpdatesPage++; renderServiceUpdates(); });
      }
      elements.serviceStartDateFilter?.addEventListener('change', () => { serviceUpdatesPage = 1; renderServiceUpdates(); });
      elements.serviceEndDateFilter?.addEventListener('change', () => { serviceUpdatesPage = 1; renderServiceUpdates(); });
      elements.serviceProviderFilter?.addEventListener('change', () => { serviceUpdatesPage = 1; renderServiceUpdates(); });
    }

    attachListeners();
    renderServiceUpdates();

    window.deleteServiceUpdate = deleteServiceUpdate;

    return {
      setServiceUpdates,
      getServiceUpdates,
      renderServiceUpdates,
      setPage(n) { serviceUpdatesPage = n; },
    };
  }

  window.ServiceUpdatesFeature = { init };
})(window);
