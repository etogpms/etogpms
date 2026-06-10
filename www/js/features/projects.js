// Projects feature module: handles CRUD, rendering, pagination, and view/edit modals.
// Depends on AppUtils (fmtNum, normalizeYmd), ProjectService, and Firebase auth.
(function (window) {
  if (window.ProjectsFeature) return;

  function init(opts) {
    const {
      elements,
      projectService,
      auth,
      utils,
      compressImage,
      elevatedAccessRef = () => false,
      isAdminRef = () => false,
    } = opts || {};
    if (!elements || !projectService || !auth || !utils) {
      throw new Error('ProjectsFeature.init missing required dependencies');
    }
    const { fmtNum, normalizeYmd, formatDateUI } = utils;
    let projects = [];
    let projectsPage = 1;

    function setProjects(list) {
      projects = Array.isArray(list) ? list.slice() : [];
    }
    function getProjects() {
      return projects.slice();
    }
    function getProjectStatus(p) {
      const accomplishments = p.accomplishments || [];
      const latest = accomplishments.length > 0
        ? accomplishments.slice().sort((a, b) => new Date(b.date) - new Date(a.date))[0]
        : null;
      if (latest) {
        if ((latest.percent ?? 0) >= 100) return 'Completed';
        if ((latest.percent ?? 0) < (latest.plannedPercent ?? 0)) return 'Delayed';
      }
      const today = new Date().toISOString().split('T')[0];
      if (p.revisedCompletion && today > p.revisedCompletion) return 'Delayed';
      if (!p.revisedCompletion && today > p.originalCompletion) return 'Delayed';
      return 'On-going';
    }
    function formatPeso(value) {
      return `₱${Number(value || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }
    function formatPercentValue(value) {
      return `${Number(value || 0).toFixed(2)}%`;
    }
    function detailField(label, value, icon = 'fa-regular fa-circle') {
      const displayValue = value || '<span class="text-muted">Not specified</span>';
      return `<div class="project-detail-field">
        <div class="project-detail-icon"><i class="${icon}"></i></div>
        <div>
          <div class="project-detail-label">${label}</div>
          <div class="project-detail-value">${displayValue}</div>
        </div>
      </div>`;
    }
    function statusBadge(status) {
      const cls = status === 'Delayed' ? 'is-delayed' : status === 'Completed' ? 'is-completed' : 'is-ongoing';
      return `<span class="project-status-badge ${cls}">${status}</span>`;
    }
    function sectionTitle(icon, title, subtitle = '') {
      return `<div class="project-section-heading">
        <div>
          <h6><i class="${icon} me-2"></i>${title}</h6>
          ${subtitle ? `<p>${subtitle}</p>` : ''}
        </div>
      </div>`;
    }

    function createBillingRow(data = {}) {
      const billingContainer = document.getElementById('billingContainer');
      if (!billingContainer) return;
      const div = document.createElement('div');
      div.className = 'row g-2 align-items-end mb-2 billing-row';
      div.innerHTML = `<div class="col-4"><input type="date" class="form-control billing-date" value="${data.date || ''}" placeholder="Date"/></div>
                  <div class="col-4"><input type="number" class="form-control billing-amount" value="${data.amount || ''}" placeholder="Amount (PHP)" step="0.01" min="0"/></div>
                  <div class="col-3"><input type="text" class="form-control billing-desc" value="${data.desc || ''}" placeholder="Description"/></div>
                  <div class="col-1 text-end"><button type="button" class="btn btn-outline-danger btn-sm remove-billing"><i class="fa fa-minus"></i></button></div>`;
      div.querySelector('.remove-billing').onclick = () => {
        div.remove();
      };
      billingContainer.appendChild(div);
    }

    function gatherBilling() {
      const billingContainer = document.getElementById('billingContainer');
      if (!billingContainer) return [];
      const rows = Array.from(billingContainer.querySelectorAll('.billing-row'));
      return rows.map(r => ({
        date: r.querySelector('.billing-date').value,
        amount: parseFloat(r.querySelector('.billing-amount').value) || 0,
        desc: r.querySelector('.billing-desc').value.trim()
      })).filter(b => b.date && b.amount);
    }
    function populateBilling(arr) {
      const billingContainer = document.getElementById('billingContainer');
      if (!billingContainer) return;
      billingContainer.innerHTML = '';
      (arr || []).forEach(d => createBillingRow(d));
    }

    async function saveProject(project) {
      await projectService.save(project);
    }

    function clearForm() {
      const billingContainer = document.getElementById('billingContainer');
      if (billingContainer) billingContainer.innerHTML = '';
      if (elements.projectForm) elements.projectForm.reset();
      const idEl = document.getElementById("projectId"); if (idEl) idEl.value = "";
      ["projectPhoto1", "projectPhoto2", "projectPhoto3"].forEach(id => { const el = document.getElementById(id); if (el) el.value = ""; });
      const linkInput = document.getElementById("contractDocsLink");
      const linkGroup = document.getElementById("contractDocsGroup");
      if (linkGroup) linkGroup.style.display = elevatedAccessRef() ? '' : 'none';
      if (linkInput) {
        linkInput.value = "";
        linkInput.disabled = !elevatedAccessRef();
      }
      const fields = ["actionTaken", "percentAccomplishment", "percentPrevious", "percentPlanned", "accompDate"];
      fields.forEach(id => { const el = document.getElementById(id); if (el) el.value = (id.includes('percent') ? 0 : ""); });
    }

    function startCreate() {
      clearForm();
      createBillingRow();
      elements.projectModal?.show();
    }

    function populateForm(project) {
      populateBilling(project.progressBilling);
      const setVal = (id, v) => { const el = document.getElementById(id); if (el) el.value = v ?? ''; };
      setVal("projectId", project.id);
      setVal("projectName", project.name);
      setVal("implementingAgency", project.implementingAgency);
      setVal("projectLocation", project.location || '');
      setVal("contractor", project.contractor);
      setVal("contractAmount", project.contractAmount);
      setVal("revisedContractAmount", project.revisedContractAmount ?? '');
      const linkGroup = document.getElementById("contractDocsGroup");
      const linkInput = document.getElementById("contractDocsLink");
      if (linkGroup) linkGroup.style.display = elevatedAccessRef() ? '' : 'none';
      if (linkInput) {
        linkInput.value = project.contractDocsLink || '';
        linkInput.disabled = !elevatedAccessRef();
      }
      setVal("ntpDate", project.ntpDate);
      setVal("originalDuration", project.originalDuration);
      setVal("timeExtension", project.timeExtension);
      setVal("originalCompletion", project.originalCompletion);
      setVal("revisedCompletion", project.revisedCompletion);
      setVal("activities", project.activities);
      setVal("issues", project.issues);
      setVal("actionTaken", "");
      setVal("remarks", project.remarks);
      setVal("otherDetails", project.otherDetails);
      const last = (project.accomplishments || []).slice(-1)[0] || { percent: 0, prevPercent: 0, date: "" };
      setVal("percentPrevious", last.prevPercent ?? last.percent ?? 0);
      setVal("percentPlanned", last.plannedPercent ?? 0);
      setVal("percentAccomplishment", last.percent ?? 0);
      setVal("accompDate", last.date);
      setVal("actionTaken", last.action || "");
    }

    function projectRowHtml(p) {
      const accomplishments = p.accomplishments || [];
      const latest = accomplishments.length > 0
        ? accomplishments.slice().sort((a, b) => new Date(b.date) - new Date(a.date))[0]
        : { percent: 0 };
      const curr = latest.percent ?? 0;
      const isAdmin = isAdminRef();
      const actionsAllowed = !isViewOnly() && auth.currentUser;
      let actionsHtml = actionsAllowed ? `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="editProject('${p.id}')"><i class="fa fa-pencil"></i></button>` : '';
      if (isAdmin) {
        actionsHtml = `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="editProject('${p.id}')"><i class="fa fa-pencil"></i></button><button class="btn btn-sm btn-danger" title="Delete" onclick="deleteProject('${p.id}')"><i class="fa fa-trash"></i></button>`;
      }
      const status = getProjectStatus(p);
      return `<tr data-id="${p.id}"><td>${p.name}</td><td>${p.implementingAgency || ''}</td><td>${p.contractor}</td><td><span class="badge bg-${status === 'Delayed' ? 'danger' : 'primary'}">${status}</span></td><td>${formatDateUI(p.revisedCompletion) || formatDateUI(p.originalCompletion)}</td><td>${curr}%</td><td class="d-flex gap-1">${actionsHtml}</td></tr>`;
    }

    function renderProjects() {
      const addBtn = document.getElementById('addProjectBtn'); if (addBtn) addBtn.style.display = (!isViewOnly() && auth.currentUser) ? 'inline-block' : 'none';
      const agencies = [...new Set(projects.map(p => p.implementingAgency))];
      const prevAgency = elements.agencyFilter?.value;
      if (elements.agencyFilter && elements.agencyFilter.options.length - 1 !== agencies.length) {
        elements.agencyFilter.innerHTML = `<option value="">All Agencies</option>` + agencies.map(a => `<option value="${a}">${a}</option>`).join("");
      }
      if (prevAgency && agencies.includes(prevAgency)) {
        elements.agencyFilter.value = prevAgency;
      }
      const text = (elements.searchInput?.value || '').toLowerCase();
      const agency = elements.agencyFilter?.value || '';
      const statusFilter = elements.statusFilter?.value || '';
      const filtered = projects.filter(p => {
        const projStatus = getProjectStatus(p);
        const matchText = (p.name || "").toLowerCase().includes(text) || (p.contractor || "").toLowerCase().includes(text);
        const matchAgency = agency ? p.implementingAgency === agency : true;
        let matchStatus = true;
        if (statusFilter) {
          if (statusFilter === 'On-going') {
            matchStatus = projStatus === 'On-going' || projStatus === 'Delayed';
          } else {
            matchStatus = projStatus === statusFilter;
          }
        }
        return matchText && matchAgency && matchStatus;
      });
      const total = filtered.length;
      const totalPages = Math.max(1, Math.ceil(total / (window.AppConstants?.PROJECTS_PER_PAGE || 30)));
      if (projectsPage > totalPages) projectsPage = totalPages;
      if (projectsPage < 1) projectsPage = 1;
      const start = (projectsPage - 1) * (window.AppConstants?.PROJECTS_PER_PAGE || 30);
      const end = start + (window.AppConstants?.PROJECTS_PER_PAGE || 30);
      const pageItems = filtered.slice(start, end);
      const tbody = document.getElementById('projectsTbody');
      if (tbody) {
        tbody.innerHTML = pageItems.map(projectRowHtml).join('');
        attachRowEvents();
      }
      buildProjectsPagination(total, projectsPage, (window.AppConstants?.PROJECTS_PER_PAGE || 30));
    }

    function buildProjectsPagination(total, page, perPage) {
      const pager = elements.projectsPagination;
      if (!pager) return;
      const totalPages = Math.max(1, Math.ceil(total / perPage));
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
      const tbody = document.getElementById('projectsTbody');
      if (!tbody) return;
      Array.from(tbody.querySelectorAll('tr')).forEach(row => {
        row.addEventListener('click', e => {
          if (isViewOnly()) return;
          if (e.target.closest('button')) return;
          const pid = row.dataset.id;
          viewProject(pid);
        });
      });
    }

    async function onSaveProject(e) {
      e.preventDefault();
      const formData = new FormData(elements.projectForm);
      const progressBilling = gatherBilling();
      const id = formData.get("projectId") || (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).substr(2, 5));
      if (!formData.get("projectName").trim()) { alert('Project Name is required'); return; }
      if (!formData.get("contractor").trim()) { alert('Contractor is required'); return; }
      const accDateRaw = formData.get("accompDate");
      const accDate = accDateRaw || new Date().toISOString().slice(0, 10);
      const currentPercent = parseFloat(formData.get("percentAccomplishment")) || 0;
      const prevPercent = parseFloat(formData.get("percentPrevious")) || 0;
      const plannedPercent = parseFloat(formData.get("percentPlanned")) || 0;
      const variance = +(currentPercent - plannedPercent).toFixed(2);
      const newAcc = { date: accDate, percent: currentPercent, prevPercent, plannedPercent, variance, activities: formData.get("activities"), issue: formData.get("issues"), action: formData.get("actionTaken"), remarks: formData.get("remarks") };
      const existing = projects.find(p => p.id === id) || {};
      const project = {
        id,
        name: formData.get("projectName"),
        implementingAgency: formData.get("implementingAgency"),
        location: formData.get("projectLocation"),
        contractor: formData.get("contractor"),
        contractAmount: parseFloat(formData.get("contractAmount")) || 0,
        revisedContractAmount: formData.get("revisedContractAmount") ? parseFloat(formData.get("revisedContractAmount")) : null,
        contractDocsLink: elevatedAccessRef() ? (formData.get("contractDocsLink")?.trim() || '') : (existing.contractDocsLink || ''),
        ntpDate: formData.get("ntpDate"),
        originalDuration: parseInt(formData.get("originalDuration"), 10) || 0,
        timeExtension: parseInt(formData.get("timeExtension"), 10) || 0,
        originalCompletion: formData.get("originalCompletion"),
        revisedCompletion: formData.get("revisedCompletion"),
        activities: formData.get("activities"),
        issues: formData.get("issues"),
        remarks: formData.get("remarks"),
        otherDetails: formData.get("otherDetails"),
        progressBilling,
        history: [...(existing.history || [])],
        photos: existing.photos || (existing.sCurveDataUrl ? [existing.sCurveDataUrl] : []),
        accomplishments: [...(existing.accomplishments || [])]
      };
      const __user = auth.currentUser;
      const userEmail = __user?.email || 'unknown';
      const userFullName = (__user?.displayName || '').trim();
      if (!project.history) project.history = [];
      project.history.push({ email: userEmail, fullName: userFullName, timestamp: new Date().toISOString(), action: formData.get("projectId") ? 'edit' : 'create' });
      if (newAcc.date) {
        const idx = project.accomplishments.findIndex(a => a.date === newAcc.date);
        if (idx > -1) {
          project.accomplishments[idx] = { ...newAcc };
        } else {
          project.accomplishments.push({ ...newAcc });
        }
      }
      function postSaveUI() {
        const idxLocal = projects.findIndex(p => p.id === project.id);
        if (idxLocal > -1) {
          projects[idxLocal] = { ...project };
        } else {
          projects.push({ ...project });
        }
        renderProjects();
        elements.searchInput && (elements.searchInput.value = '');
        elements.agencyFilter && (elements.agencyFilter.value = '');
        elements.statusFilter && (elements.statusFilter.value = '');
        elements.projectModal?.hide();
        clearForm();
        const billingContainer = document.getElementById('billingContainer');
        if (billingContainer) billingContainer.innerHTML = '';
      }
      const photoInputs = ["projectPhoto1", "projectPhoto2", "projectPhoto3"].map(id => document.getElementById(id));
      const files = photoInputs.map(inp => inp?.files[0]).filter(f => !!f);
      if (files.length && typeof compressImage === 'function') {
        Promise.all(files.map(f => compressImage(f, 1024, 0.75)))
          .then(dataUrls => {
            project.photos = dataUrls.slice(0, 3);
            return saveProject(project);
          })
          .then(postSaveUI)
          .catch(err => alert(err.message));
      } else {
        saveProject(project).then(postSaveUI).catch(err => alert(err.message));
      }
    }

    async function deleteProject(id) {
      if (!elevatedAccessRef()) { alert('Only admin/level2 can delete projects'); return; }
      try {
        await projectService.remove(id);
        projects = projects.filter(p => p.id !== id);
        renderProjects();
      } catch (err) { alert(err.message); }
    }

    function attachFormListeners() {
      if (elements.projectForm) elements.projectForm.addEventListener('submit', onSaveProject);
      if (elements.searchInput) elements.searchInput.addEventListener('input', () => { projectsPage = 1; renderProjects(); });
      elements.agencyFilter?.addEventListener('change', () => { projectsPage = 1; renderProjects(); });
      elements.statusFilter?.addEventListener('change', () => { projectsPage = 1; renderProjects(); });
      const addBillingBtn = document.getElementById('addBillingBtn');
      if (addBillingBtn) addBillingBtn.addEventListener('click', () => createBillingRow());
      const addProjectBtn = document.getElementById('addProjectBtn');
      if (addProjectBtn) addProjectBtn.addEventListener('click', () => startCreate());
      if (elements.projectsPagination) {
        elements.projectsPagination.addEventListener('click', (e) => {
          if (e.target.tagName.toLowerCase() !== 'button') return;
          const page = e.target.getAttribute('data-page');
          if (!page) return;
          if (page === 'prev') {
            if (projectsPage > 1) projectsPage--;
          } else if (page === 'next') {
            projectsPage++;
          } else {
            const n = parseInt(page, 10);
            if (!isNaN(n)) projectsPage = n;
          }
          renderProjects();
          window.scrollTo({ top: 0, behavior: 'smooth' });
        });
      }
    }

    function isViewOnly() {
      if (typeof window.__VIEW_ONLY__ === 'function') return !!window.__VIEW_ONLY__();
      return !!window.__VIEW_ONLY__ || false;
    }

    window.deleteProject = deleteProject;
    window.startProjectCreate = startCreate;
    window.editProject = (id) => { const p = projects.find(x => x.id === id); if (!p) return; populateForm(p); elements.projectModal?.show(); };
    window.viewProject = (id) => {
      const p = projects.find(proj => proj.id === id); if (!p) return;
      window.__prCurrentItem = p; // Store for export
      const photos = (p.photos && p.photos.length) ? p.photos : (p.sCurveDataUrl ? [p.sCurveDataUrl] : []);
      const billingHtml = (p.progressBilling && p.progressBilling.length) ? `<section class="project-detail-section">${sectionTitle('fa-solid fa-receipt', 'Billing Details', 'Recorded progress billing transactions.')}<div class="table-responsive project-table-wrap"><table class="table table-sm project-record-table"><thead><tr><th>Date</th><th>Amount (PHP)</th><th>Description</th></tr></thead><tbody>${p.progressBilling.map(b => `<tr><td>${formatDateUI(b.date)}</td><td class="text-nowrap fw-semibold">${formatPeso(b.amount)}</td><td>${b.desc || ''}</td></tr>`).join('')}</tbody></table></div></section>` : '';
      const gridCls = (photos.length <= 2) ? 'photo-grid single mb-3' : 'photo-grid mb-3';
      const photosHtml = photos.length ? `<section class="project-detail-section">${sectionTitle('fa-regular fa-image', 'Project Documentation', 'Latest uploaded project photos and visual records.')}<div class="${gridCls}">` + photos.slice(0, 3).map((url, i) => `<div class="photo-tile" aria-label="Project photo ${i + 1}"><img src="${url}" data-full="${url}" data-index="${i}" alt="Project photo ${i + 1}" class="photo-img" loading="lazy"></div>`).join('') + `</div></section>` : '';
      const accompSorted = (p.accomplishments || []).slice().sort((a, b) => new Date(b.date) - new Date(a.date));
      const seenKeys = new Set();
      const accompDedup = accompSorted.filter(a => {
        const key = `${a.date}|${a.percent}|${a.prevPercent}|${a.plannedPercent}|${a.activities || ''}|${a.issue || ''}|${a.action || ''}|${a.remarks || ''}`;
        if (seenKeys.has(key)) return false;
        seenKeys.add(key);
        return true;
      });
      const editHistorySorted = (p.history || []).slice().sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      const body = document.getElementById('detailsBody');
      if (!body) return;
      const latest = accompDedup[0] || {};
      const status = getProjectStatus(p);
      const targetCompletion = formatDateUI(p.revisedCompletion) || formatDateUI(p.originalCompletion);
      const duration = `${p.originalDuration || 0} days${p.timeExtension ? ` + ${p.timeExtension}` : ''}`;
      const docsHtml = p.contractDocsLink ? (elevatedAccessRef() ? `<a class="btn btn-sm btn-outline-primary project-docs-link" href="${p.contractDocsLink}" target="_blank" rel="noopener"><i class="fa-regular fa-folder-open me-1"></i>Open Contract Documents</a>` : `<span class="text-muted">No authority to access</span>`) : '<span class="text-muted">No document link provided</span>';
      const accomplishRows = accompDedup.length ? accompDedup.map(a => `<tr>
        <td class="date-col">${formatDateUI(a.date)}</td>
        <td class="percent-col">${formatPercentValue(a.plannedPercent)}</td>
        <td class="percent-col">${formatPercentValue(a.prevPercent)}</td>
        <td class="percent-col fw-semibold">${formatPercentValue(a.percent)}</td>
        <td class="percent-col ${Number(a.variance || 0) < 0 ? 'project-variance-bad' : 'project-variance-good'}">${Number(a.variance || 0) >= 0 ? '+' : ''}${formatPercentValue(a.variance)}</td>
        <td>${(typeof bulletizeActivities === 'function') ? bulletizeActivities(a.activities) : (a.activities || '')}</td>
        <td>${a.issue || '<span class="text-muted">n/a</span>'}</td>
        <td>${a.action || ''}</td>
        <td>${a.remarks || ''}</td>
      </tr>`).join('') : `<tr><td colspan="9" class="text-center text-muted py-4">No accomplishment entries recorded.</td></tr>`;
      const editHistoryHtml = (isAdminRef() && editHistorySorted.length) ? `<section class="project-detail-section">${sectionTitle('fa-regular fa-clock', 'Edit History', 'Administrative activity log.')}<div class="project-audit-list">${editHistorySorted.map(h => {
        const when = h.timestamp ? new Date(h.timestamp).toLocaleString() : '';
        const who = h.fullName || h.email || 'Unknown user';
        const act = h.action || 'update';
        return `<div class="project-audit-item">
          <i class="fa-regular fa-clock"></i>
          <div>
            <div><strong>${act}</strong> by ${who}</div>
            <time>${when}</time>
          </div>
        </div>`;
      }).join('')}</div></section>` : '';
      body.innerHTML = `<div class="project-record">
        <div class="project-record-hero">
          <div class="project-record-seal"><i class="fa-solid fa-building-columns"></i></div>
          <div class="project-record-title">
            <div class="project-record-kicker">MWSS Project Monitoring Record</div>
            <h5>${p.name || 'Untitled Project'}</h5>
            <p>${p.implementingAgency || 'Implementing agency not specified'}</p>
          </div>
          <div class="project-record-status">${statusBadge(status)}</div>
        </div>

        <div class="project-summary-grid">
          <div class="project-summary-metric">
            <span>Contract Amount</span>
            <strong>${formatPeso(p.contractAmount)}</strong>
          </div>
          ${p.revisedContractAmount ? `<div class="project-summary-metric">
            <span>Revised Amount</span>
            <strong>${formatPeso(p.revisedContractAmount)}</strong>
          </div>` : ''}
          <div class="project-summary-metric">
            <span>Physical Accomplishment</span>
            <strong>${formatPercentValue(latest.percent)}</strong>
          </div>
          <div class="project-summary-metric">
            <span>Target Completion</span>
            <strong>${targetCompletion || 'Not specified'}</strong>
          </div>
        </div>

        <section class="project-detail-section">
          ${sectionTitle('fa-regular fa-file-lines', 'Project Information', 'Official contract and implementation details.')}
          <div class="project-detail-grid">
            ${detailField('Contractor', p.contractor, 'fa-regular fa-building')}
            ${detailField('Location', p.location, 'fa-solid fa-location-dot')}
            ${detailField('Notice to Proceed', formatDateUI(p.ntpDate), 'fa-regular fa-calendar-check')}
            ${detailField('Duration', duration, 'fa-regular fa-hourglass-half')}
            ${detailField('Contract Documents', docsHtml, 'fa-regular fa-folder-open')}
            ${p.otherDetails ? detailField('Other Details', p.otherDetails, 'fa-regular fa-note-sticky') : ''}
          </div>
        </section>

        <section class="project-detail-section">
          ${sectionTitle('fa-solid fa-chart-line', 'Accomplishment History', 'Progress updates, issues, actions taken, and remarks.')}
          <div class="table-responsive project-table-wrap">
            <table class="table table-sm project-record-table project-accomplishment-table">
              <thead>
                <tr>
                  <th>Date</th>
                  <th class="percent-col">Planned %</th>
                  <th class="percent-col">Previous %</th>
                  <th class="percent-col">To Date %</th>
                  <th class="percent-col">Variance %</th>
                  <th>Activities</th>
                  <th>Issue</th>
                  <th>Action Taken</th>
                  <th>Remarks</th>
                </tr>
              </thead>
              <tbody>${accomplishRows}</tbody>
            </table>
          </div>
        </section>

        ${billingHtml}
        ${photosHtml}
        ${editHistoryHtml}
      </div>`;
      document.getElementById('detailsModal')?.classList.add('show');
      const modal = window.bootstrap ? window.bootstrap.Modal.getOrCreateInstance(document.getElementById('detailsModal')) : null;
      modal && modal.show();
    };

    attachFormListeners();

    function exportWord() {
      if (window.exportProjectDocx && window.__prCurrentItem) {
        window.exportProjectDocx(window.__prCurrentItem, { open: true })
          .catch(e => {
            console.error(e);
            alert('Export failed: ' + e.message);
          });
        return;
      }
      // Fallback if export.js not loaded
      if (window.prGenerateFromTemplate && window.__prCurrentItem) {
        window.prGenerateFromTemplate(window.__prCurrentItem);
        return;
      }

      const content = document.getElementById('detailsBody').innerHTML;
      if (!content || !window.htmlDocx || !window.saveAs) {
        alert('Export module not loaded or content empty.');
        return;
      }
      // Simple HTML wrapper for styling
      const fullHtml = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <style>
            body { font-family: 'Calibri', sans-serif; font-size: 11pt; }
            h5 { font-size: 14pt; font-weight: bold; margin-bottom: 10px; }
            h6 { font-size: 12pt; font-weight: bold; margin-top: 15px; margin-bottom: 5px; }
            p { margin: 5px 0; }
            table { border-collapse: collapse; width: 100%; margin-top: 10px; }
            th, td { border: 1px solid #999; padding: 4px; font-size: 10pt; text-align: left; }
            th { background-color: #eee; font-weight: bold; }
            .badge { display: inline-block; padding: 2px 6px; border: 1px solid #ccc; border-radius: 4px; }
          </style>
        </head>
        <body>
          ${content}
        </body>
        </html>
      `;
      try {
        const blob = window.htmlDocx.asBlob(fullHtml);
        const title = (document.querySelector('#detailsBody h5')?.innerText || 'ProjectDetails').replace(/[\\/:*?"<>|]/g, '_');
        window.saveAs(blob, `${title}.docx`);
      } catch (e) {
        console.error(e);
        alert('Export failed: ' + e.message);
      }
      alert('Using legacy HTML export.');
      // (Simplified fallback code intentionally omitted to rely on above)
    }

    // Expose exportWord on the feature object
    window.ProjectsFeature = window.ProjectsFeature || {};
    window.ProjectsFeature.exportWord = exportWord;

    return {
      setProjects,
      getProjects,
      renderProjects,
      getProjectStatus,
      setPage(n) { projectsPage = n; },
      exportWord
    };
  }

  window.ProjectsFeature = { init };
})(window);
