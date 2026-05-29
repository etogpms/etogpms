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
      const billingHtml = (p.progressBilling && p.progressBilling.length) ? `<h6 class="mt-3">Billing Details</h6><div class="table-responsive"><table class="table table-sm"><thead><tr><th>Date</th><th>Amount (PHP)</th><th>Description</th></tr></thead><tbody>${p.progressBilling.map(b => `<tr><td>${formatDateUI(b.date)}</td><td>₱${Number(b.amount).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td><td>${b.desc || ''}</td></tr>`).join('')}</tbody></table></div>` : '';
      const gridCls = (photos.length <= 2) ? 'photo-grid single mb-3' : 'photo-grid mb-3';
      const photosHtml = photos.length ? `<div class="${gridCls}">` + photos.slice(0, 3).map((url, i) => `<div class="photo-tile" aria-label="Project photo ${i + 1}"><img src="${url}" data-full="${url}" data-index="${i}" alt="Project photo ${i + 1}" class="photo-img" loading="lazy"></div>`).join('') + `</div>` : '';
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
      body.innerHTML = `<h5 class="fw-bold mb-2">${p.name}</h5><p><strong>Implementing Agency:</strong> ${p.implementingAgency}</p><p><strong>Contractor:</strong> ${p.contractor}</p>
<p><strong>Location:</strong> ${p.location || ''}</p><p><strong>Contract Amount:</strong> ₱${Number(p.contractAmount || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>${p.revisedContractAmount ? `<p><strong>Revised Contract Amount:</strong> ₱${Number(p.revisedContractAmount).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>` : ''}<p><strong>Status:</strong> ${getProjectStatus(p)}</p><p><strong>NTP:</strong> ${formatDateUI(p.ntpDate)}</p><p><strong>Duration:</strong> ${p.originalDuration} days ${p.timeExtension ? "+" + p.timeExtension : ""}</p><p><strong>Target Completion:</strong> ${formatDateUI(p.revisedCompletion) || formatDateUI(p.originalCompletion)}</p>${p.otherDetails ? `<p><strong>Details:</strong> ${p.otherDetails}</p>` : ''}${p.contractDocsLink ? `<p><strong>Contract Docs:</strong> ${elevatedAccessRef() ? `<a href="${p.contractDocsLink}" target="_blank">Open</a>` : `<span class="text-muted">No authority to access</span>`}</p>` : ''}<h6 class="mt-3">Accomplishment History</h6><div class="table-responsive"><table class="table table-sm"><thead><tr><th>Date</th><th class="percent-col">Planned %</th><th class="percent-col">Previous %</th><th class="percent-col">To Date %</th><th class="percent-col">Variance %</th><th class="w-25">Activities</th><th class="w-20">Issue</th><th class="w-20">Action Taken</th><th class="w-25">Remarks</th></tr></thead><tbody>${accompDedup.map(a => `<tr><td class="date-col">${formatDateUI(a.date)}</td><td class="percent-col">${(a.plannedPercent ?? 0).toFixed(2)}%</td><td class="percent-col">${(a.prevPercent ?? 0).toFixed(2)}%</td><td class="percent-col">${(a.percent ?? 0).toFixed(2)}%</td><td class="percent-col">${a.variance >= 0 ? '+' : ''}${(a.variance ?? 0).toFixed(2)}%</td><td class="w-25">${(typeof bulletizeActivities === 'function') ? bulletizeActivities(a.activities) : (a.activities || '')}</td><td class="w-20">${a.issue || ''}</td><td class="w-20">${a.action || ''}</td><td class="w-25">${a.remarks || ''}</td></tr>`).join('')}</tbody></table></div>${billingHtml}${photosHtml}${editHistorySorted.length ? `<h6 class="mt-3">Edit History</h6><div class="table-responsive"><table class="table table-sm"><thead><tr><th>User</th><th>Timestamp</th><th>Action</th></tr></thead><tbody>${editHistorySorted.map(h => `<tr><td>${h.email}</td><td>${new Date(h.timestamp).toLocaleString()}</td><td>${h.action}</td></tr>`).join('')}</tbody></table></div>` : ''}`;
      try {
        const histHdr = Array.from(body.querySelectorAll('h6')).find(h => /edit history/i.test(h.textContent || ''));
        const histWrap = histHdr ? histHdr.nextElementSibling : null;
        if (histHdr && histWrap) {
          if (!isAdminRef()) {
            histHdr.remove();
            histWrap.remove();
          } else {
            const rows = editHistorySorted.map(h => {
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
            histHdr.outerHTML = `<div class="small text-muted mt-3">Edit History</div>`;
            const container = document.createElement('div');
            container.innerHTML = `<div class="border rounded p-2 bg-light small">${rows}</div>`;
            histWrap.replaceWith(container);
            setTimeout(() => {
              const photosGrid = body.querySelector('.photo-grid');
              const histLabel = Array.from(body.querySelectorAll('div.small.text-muted')).find(d => /edit history/i.test(d.textContent || ''));
              if (photosGrid && histLabel) {
                const histContainer = histLabel.nextElementSibling;
                if (histContainer) {
                  photosGrid.insertAdjacentElement('afterend', histLabel);
                  histLabel.insertAdjacentElement('afterend', histContainer);
                }
              }
            }, 0);
          }
        }
      } catch (_) { /* noop */ }
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
