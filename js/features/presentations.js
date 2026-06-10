// Product Presentations feature module: handles CRUD, rendering, filtering, and modal logic.
// Depends on PresentationService, Firebase auth, DOM elements, and AppUtils.
(function (window) {
  if (window.PresentationsFeature) return;

  function init(opts) {
    const {
      elements,
      presentationService,
      auth,
      utils,
      calendarService,
    } = opts || {};
    if (!elements || !presentationService || !auth || !utils) {
      throw new Error('PresentationsFeature.init missing required dependencies');
    }
    const { normalizeYmd, todayYmd, jan1OfYear, dec31OfYear, debounce } = utils;
    let presentations = [];

    const isViewOnly = () => {
      if (typeof window.__VIEW_ONLY__ === 'function') return !!window.__VIEW_ONLY__();
      return !!window.__VIEW_ONLY__ || false;
    };
    const isAdminRef = () => {
      if (typeof opts.isAdminRef === 'function') return opts.isAdminRef();
      return false;
    };
    const elevatedAccessRef = () => {
      if (typeof opts.elevatedAccessRef === 'function') return opts.elevatedAccessRef();
      return false;
    };
    const escapeHtml = (value) => {
      if (window.AppUtils && window.AppUtils.escapeHtml) return window.AppUtils.escapeHtml(value);
      return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    };

    function setPresentations(list) {
      presentations = Array.isArray(list) ? list.slice() : [];
    }
    function getPresentations() {
      return presentations.slice();
    }

    // ---- Filtering ----
    function getFilteredPresentations() {
      let list = presentations.slice();
      const startDate = elements.presentationStartDateFilter?.value || '';
      const endDate = elements.presentationEndDateFilter?.value || '';
      const subjectQ = (elements.presentationSubjectFilter?.value || '').trim().toLowerCase();
      const presenterQ = (elements.presentationPresenterFilter?.value || '').trim().toLowerCase();

      if (startDate) {
        list = list.filter(p => {
          const d = normalizeYmd(p.date);
          return d && d >= startDate;
        });
      }
      if (endDate) {
        list = list.filter(p => {
          const d = normalizeYmd(p.date);
          return d && d <= endDate;
        });
      }
      if (subjectQ) {
        list = list.filter(p => (p.subject || '').toLowerCase().includes(subjectQ));
      }
      if (presenterQ) {
        list = list.filter(p => (p.presenter || '').toLowerCase().includes(presenterQ));
      }

      // Sort by date descending (most recent first)
      list.sort((a, b) => {
        const da = normalizeYmd(a.date) || '';
        const db = normalizeYmd(b.date) || '';
        return db.localeCompare(da);
      });

      return list;
    }

    // ---- Row HTML ----
    function presentationRowHtml(p) {
      const canEdit = !isViewOnly() && auth.currentUser;
      let actionsHtml = '';
      if (canEdit) {
        actionsHtml += `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="event.stopPropagation(); editPresentation('${p.id}')"><i class="fa fa-pencil"></i></button>`;
      }
      if (isAdminRef() || elevatedAccessRef()) {
        actionsHtml += `<button class="btn btn-sm btn-danger" title="Delete" onclick="event.stopPropagation(); deletePresentation('${p.id}')"><i class="fa fa-trash"></i></button>`;
      }
      const dateStr = window.AppUtils?.formatDateUI ? window.AppUtils.formatDateUI(p.date) : (p.date || '');
      let timeStr = p.time || '';
      if (timeStr) {
        const [h, m] = timeStr.split(':');
        const hour = parseInt(h, 10);
        const ampm = hour >= 12 ? 'PM' : 'AM';
        const h12 = hour % 12 || 12;
        timeStr = `${h12}:${m} ${ampm}`;
      }
      return `<tr data-id="${p.id}" class="pres-clickable-row" onclick="window.showPresentationDetails('${p.id}')" style="cursor:pointer;">
        <td class="text-nowrap">${dateStr}</td>
        <td class="text-start">${p.subject || ''}</td>
        <td class="text-nowrap">${timeStr}</td>
        <td>${p.venue || ''}</td>
        <td>${p.presenter || ''}</td>
        <td class="d-flex gap-1 justify-content-center">${actionsHtml}</td>
      </tr>`;
    }

    function presentationCardHtml(p) {
      const canEdit = !isViewOnly() && auth.currentUser;
      let actionsHtml = '';
      if (canEdit) {
        actionsHtml += `<button class="btn btn-sm btn-outline-primary me-2" title="Edit" onclick="event.stopPropagation(); editPresentation('${p.id}')"><i class="fa fa-pencil"></i> Edit</button>`;
      }
      if (isAdminRef() || elevatedAccessRef()) {
        actionsHtml += `<button class="btn btn-sm btn-outline-danger" title="Delete" onclick="event.stopPropagation(); deletePresentation('${p.id}')"><i class="fa fa-trash"></i> Delete</button>`;
      }
      const dateStr = window.AppUtils?.formatDateUI ? window.AppUtils.formatDateUI(p.date) : (p.date || '');
      let timeStr = p.time || '';
      if (timeStr) {
        const [h, m] = timeStr.split(':');
        const hour = parseInt(h, 10);
        const ampm = hour >= 12 ? 'PM' : 'AM';
        const h12 = hour % 12 || 12;
        timeStr = `${h12}:${m} ${ampm}`;
      }
      const esc = (s) => {
        if (window.AppUtils && window.AppUtils.escapeHtml) return window.AppUtils.escapeHtml(s);
        return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      };
      
      return `<div class="card mb-3 border border-light shadow-sm" onclick="window.showPresentationDetails('${p.id}')" style="cursor:pointer; border-radius: 12px; transition: transform 0.15s ease, box-shadow 0.15s ease;">
        <div class="card-body p-3">
          <div class="d-flex justify-content-between align-items-start mb-2">
            <span class="badge bg-primary-subtle text-primary border border-primary-subtle px-2 py-1 small rounded-pill">
              <i class="fa-regular fa-calendar me-1"></i>${dateStr}
            </span>
            <span class="text-muted small">
              <i class="fa-regular fa-clock me-1"></i>${timeStr || '—'}
            </span>
          </div>
          <h6 class="card-title fw-bold text-dark mb-3">${esc(p.subject || 'No Subject')}</h6>
          <div class="row g-2 mb-3 text-secondary small">
            <div class="col-6">
              <div class="d-flex align-items-center">
                <i class="fa-solid fa-location-dot text-danger me-2" style="width: 14px;"></i>
                <span class="text-truncate">${esc(p.venue || '—')}</span>
              </div>
            </div>
            <div class="col-6">
              <div class="d-flex align-items-center">
                <i class="fa-solid fa-user-tie text-primary me-2" style="width: 14px;"></i>
                <span class="text-truncate">${esc(p.presenter || '—')}</span>
              </div>
            </div>
          </div>
          ${actionsHtml ? `
            <div class="d-flex justify-content-end border-top pt-2 mt-2">
              ${actionsHtml}
            </div>
          ` : ''}
        </div>
      </div>`;
    }

    // ---- Render ----
    function renderPresentations() {
      const list = getFilteredPresentations();
      const tbody = elements.presentationsTbody;
      const mobileList = elements.presentationsMobileList;
      const emptyEl = elements.presentationsEmpty;

      // Update year label
      if (elements.presentationYearLabel) {
        const yr = elements.presentationYearFilter?.value;
        if (yr === 'ALL') {
          elements.presentationYearLabel.textContent = 'All Years';
        } else if (yr) {
          elements.presentationYearLabel.textContent = yr;
        } else {
          elements.presentationYearLabel.textContent = new Date().getFullYear();
        }
      }

      // Show/hide add button based on auth
      if (elements.addPresentationBtn) {
        elements.addPresentationBtn.style.display = (!isViewOnly() && auth.currentUser) ? 'inline-block' : 'none';
      }

      if (list.length === 0) {
        if (tbody) tbody.innerHTML = '';
        if (mobileList) mobileList.innerHTML = '';
        if (emptyEl) emptyEl.style.display = 'block';
        return;
      }
      if (emptyEl) emptyEl.style.display = 'none';
      if (tbody) {
        tbody.innerHTML = list.map(presentationRowHtml).join('');
      }
      if (mobileList) {
        mobileList.innerHTML = list.map(presentationCardHtml).join('');
      }
    }

    // ---- Presentation Details Modal ----
    let presDetailsModal = null;
    let currentDetailId = null;

    function formatPresentationTime(value) {
      let timeStr = value || '';
      if (timeStr) {
        const [h, m] = timeStr.split(':');
        const hour = parseInt(h, 10);
        if (!Number.isNaN(hour)) {
          const ampm = hour >= 12 ? 'PM' : 'AM';
          const h12 = hour % 12 || 12;
          timeStr = `${h12}:${m || '00'} ${ampm}`;
        }
      }
      return timeStr;
    }

    function formatPresentationDate(value, opts = {}) {
      if (!value) return '—';
      try {
        const [y, mon, day] = String(value).split('-').map(Number);
        const dateObj = new Date(y, mon - 1, day);
        return dateObj.toLocaleDateString('en-US', {
          weekday: opts.weekday === false ? undefined : 'short',
          year: 'numeric',
          month: 'long',
          day: 'numeric',
        });
      } catch (_) {
        return String(value);
      }
    }

    function formatPresentationStamp(value) {
      if (!value) return '';
      try {
        if (typeof value === 'string') return new Date(value).toLocaleString();
        if (value.toDate) return value.toDate().toLocaleString();
      } catch (_) { /* keep blank */ }
      return '';
    }

    function presentationStatus(dateValue) {
      const dateYmd = normalizeYmd(dateValue);
      if (!dateYmd) return { text: 'For Scheduling', className: 'is-pending' };
      const today = todayYmd();
      if (dateYmd < today) return { text: 'Completed', className: 'is-completed' };
      if (dateYmd === today) return { text: 'Today', className: 'is-today' };
      return { text: 'Scheduled', className: 'is-scheduled' };
    }

    function presValue(value, fallback = '—') {
      const text = String(value ?? '').trim();
      if (!text) return `<span class="presentation-muted">${escapeHtml(fallback)}</span>`;
      return escapeHtml(text);
    }

    function presField(label, value, icon, options = {}) {
      const content = options.html ? value : presValue(value);
      return `<div class="presentation-detail-field${options.wide ? ' is-wide' : ''}">
        <div class="presentation-detail-icon"><i class="${icon}"></i></div>
        <div>
          <div class="presentation-detail-label">${escapeHtml(label)}</div>
          <div class="presentation-detail-value">${content}</div>
        </div>
      </div>`;
    }

    function presSection(icon, title, fields, subtitle = '') {
      return `<section class="presentation-detail-section">
        <div class="presentation-section-heading">
          <div>
            <h6><i class="${icon} me-2"></i>${escapeHtml(title)}</h6>
            ${subtitle ? `<p>${escapeHtml(subtitle)}</p>` : ''}
          </div>
        </div>
        <div class="presentation-detail-grid">
          ${fields.join('')}
        </div>
      </section>`;
    }

    function presAuditSection(p) {
      const updatedAt = formatPresentationStamp(p.updatedAt);
      const items = [];
      if (p.updatedBy || updatedAt) {
        items.push(`<div class="presentation-audit-item">
          <i class="fa-regular fa-clock"></i>
          <div><strong>Last updated</strong><br><span>${presValue(p.updatedBy || 'System')}</span>${updatedAt ? `<time>${escapeHtml(updatedAt)}</time>` : ''}</div>
        </div>`);
      }
      if (!items.length) {
        return `<section class="presentation-detail-section">
          <div class="presentation-section-heading">
            <h6><i class="fa-regular fa-clock me-2"></i>Audit Information</h6>
          </div>
          <div class="presentation-empty-state">No audit information recorded.</div>
        </section>`;
      }
      return `<section class="presentation-detail-section">
        <div class="presentation-section-heading">
          <h6><i class="fa-regular fa-clock me-2"></i>Audit Information</h6>
        </div>
        <div class="presentation-audit-list">${items.join('')}</div>
      </section>`;
    }

    function initDetailsModal() {
      const el = document.getElementById('presentationDetailsModal');
      if (el) {
        presDetailsModal = new bootstrap.Modal(el);
      }

      // Edit button in details modal
      const editBtn = document.getElementById('presDetailEditBtn');
      if (editBtn) {
        editBtn.addEventListener('click', () => {
          if (presDetailsModal) presDetailsModal.hide();
          setTimeout(() => {
            if (currentDetailId) openPresentationModal(currentDetailId);
          }, 300);
        });
      }

      // Delete button in details modal
      const deleteBtn = document.getElementById('presDetailDeleteBtn');
      if (deleteBtn) {
        deleteBtn.addEventListener('click', async () => {
          if (!currentDetailId) return;
          if (presDetailsModal) presDetailsModal.hide();
          setTimeout(() => deletePresentation(currentDetailId), 300);
        });
      }
    }

    function showPresentationDetails(id) {
      const p = presentations.find(x => x.id === id);
      if (!p) return;
      currentDetailId = id;

      const body = document.getElementById('presDetailBody');
      const timeStr = formatPresentationTime(p.time) || '—';
      const dateDisplay = formatPresentationDate(p.date);
      const status = presentationStatus(p.date);
      const remarksHtml = `<div class="presentation-remarks">${presValue(p.remarks, 'No remarks recorded.')}</div>`;

      if (body) {
        body.innerHTML = `<div class="presentation-record">
          <div class="presentation-record-hero">
            <div class="presentation-record-seal"><i class="fa-solid fa-chalkboard-user"></i></div>
            <div class="presentation-record-title">
              <div class="presentation-record-kicker">MWSS Product Presentation Record</div>
              <h5>${presValue(p.subject, 'Untitled Presentation')}</h5>
              <p>${presValue(p.presenter, 'Presenter not specified')}</p>
            </div>
            <div class="presentation-record-status">
              <span class="presentation-status-badge ${status.className}">${escapeHtml(status.text)}</span>
            </div>
          </div>

          <div class="presentation-summary-grid">
            <div class="presentation-summary-metric"><span>Date</span><strong>${escapeHtml(dateDisplay)}</strong></div>
            <div class="presentation-summary-metric"><span>Time</span><strong>${escapeHtml(timeStr)}</strong></div>
            <div class="presentation-summary-metric"><span>Venue</span><strong>${presValue(p.venue)}</strong></div>
          </div>

          ${presSection('fa-regular fa-calendar-check', 'Schedule Information', [
            presField('Presentation Date', dateDisplay, 'fa-regular fa-calendar-days'),
            presField('Presentation Time', timeStr, 'fa-regular fa-clock'),
            presField('Venue', p.venue, 'fa-solid fa-location-dot', { wide: true }),
          ], 'Official schedule and venue details')}

          ${presSection('fa-solid fa-user-tie', 'Presentation Details', [
            presField('Subject', p.subject, 'fa-solid fa-file-lines', { wide: true }),
            presField('Presenter', p.presenter, 'fa-solid fa-user-tie'),
            presField('Remarks', remarksHtml, 'fa-regular fa-message', { html: true, wide: true }),
          ], 'Presenter information and notes')}

          ${presAuditSection(p)}
        </div>`;
      }

      // Show/hide Edit/Delete buttons
      const actionsEl = document.getElementById('presDetailActions');
      const canEdit = !isViewOnly() && auth.currentUser;
      const canDelete = isAdminRef() || elevatedAccessRef();
      if (actionsEl) {
        actionsEl.style.display = (canEdit || canDelete) ? 'flex' : 'none';
      }
      const editBtn = document.getElementById('presDetailEditBtn');
      if (editBtn) editBtn.style.display = canEdit ? 'inline-block' : 'none';
      const deleteBtn = document.getElementById('presDetailDeleteBtn');
      if (deleteBtn) deleteBtn.style.display = canDelete ? 'inline-block' : 'none';

      if (presDetailsModal) presDetailsModal.show();
    }

    // ---- Modal open/reset ----
    function resetPresentationForm() {
      if (elements.presentationForm) elements.presentationForm.reset();
      if (elements.presentationId) elements.presentationId.value = '';
    }

    function openPresentationModal(id) {
      resetPresentationForm();
      if (id) {
        // Edit mode: populate form
        const p = presentations.find(x => x.id === id);
        if (p) {
          if (elements.presentationId) elements.presentationId.value = p.id;
          if (elements.presentationSubject) elements.presentationSubject.value = p.subject || '';
          if (elements.presentationDate) elements.presentationDate.value = p.date || '';
          if (elements.presentationTime) elements.presentationTime.value = p.time || '';
          if (elements.presentationVenue) elements.presentationVenue.value = p.venue || '';
          if (elements.presentationPresenter) elements.presentationPresenter.value = p.presenter || '';
          if (elements.presentationRemarks) elements.presentationRemarks.value = p.remarks || '';
        }
      }
      if (elements.presentationModal && typeof elements.presentationModal.show === 'function') {
        elements.presentationModal.show();
      }
    }

    // ---- Save ----
    async function onSavePresentation(formData) {
      const id = (formData.get('presentationId') || '').trim() || (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36));
      const user = auth.currentUser;
      const userEmail = user?.email || 'viewer';

      const payload = {
        id,
        subject: (formData.get('presentationSubject') || '').trim(),
        date: formData.get('presentationDate') || '',
        time: formData.get('presentationTime') || '',
        venue: (formData.get('presentationVenue') || '').trim(),
        presenter: (formData.get('presentationPresenter') || '').trim(),
        remarks: (formData.get('presentationRemarks') || '').trim(),
        updatedBy: userEmail,
        updatedAt: new Date().toISOString(),
      };

      if (!payload.subject) { alert('Subject is required.'); return; }
      if (!payload.date) { alert('Date is required.'); return; }

      try {
        await presentationService.save(payload);

        // Smart-link: auto-create/update a Calendar event for this presentation
        await syncPresentationToCalendar(payload);

        if (elements.presentationModal && typeof elements.presentationModal.hide === 'function') {
          elements.presentationModal.hide();
        }
        resetPresentationForm();
      } catch (err) {
        console.error('[Presentations] save error', err);
        if (err && err.code === 'permission-denied') {
          alert('Missing or insufficient permissions. Please ensure Firestore rules allow creating presentations.');
        } else {
          alert(err.message || String(err));
        }
      }
    }

    // ---- Delete ----
    async function deletePresentation(id) {
      if (!confirm('Delete this presentation?')) return;
      try {
        await presentationService.remove(id);

        // Smart-link: remove the linked calendar event
        await removePresentationFromCalendar(id);
      } catch (err) {
        console.error('[Presentations] delete error', err);
        alert(err.message || String(err));
      }
    }

    // ---- Smart-link: Calendar sync helpers ----

    /**
     * Create or update a calendar event linked to a product presentation.
     * Uses the presentation ID as part of the calendar event ID for 1:1 mapping.
     */
    async function syncPresentationToCalendar(presentation) {
      if (!calendarService) return;
      const user = auth.currentUser;
      if (!user) return;

      const calEventId = 'pres_' + presentation.id;

      // Build time info
      let startTime = presentation.time || null;
      let endTime = null;
      if (startTime) {
        // Estimate 1-hour duration for the presentation
        const [h, m] = startTime.split(':').map(Number);
        const endH = (h + 1) % 24;
        endTime = `${String(endH).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
      }

      // Build a description from the presentation details
      const descParts = [];
      if (presentation.presenter) descParts.push('Presenter: ' + presentation.presenter);
      if (presentation.venue) descParts.push('Venue: ' + presentation.venue);
      if (presentation.remarks) descParts.push(presentation.remarks);

      const calPayload = {
        id: calEventId,
        title: presentation.subject || 'Product Presentation',
        type: 'meeting',
        startDate: presentation.date,
        endDate: presentation.date,
        allDay: !startTime,
        startTime: startTime,
        endTime: endTime,
        description: descParts.join('\n'),
        ownerUid: user.uid,
        ownerName: user.displayName || user.email,
        linkedPresentationId: presentation.id,
        createdAt: window.AppFirebase.serverTimestamp(),
        updatedAt: window.AppFirebase.serverTimestamp(),
      };

      try {
        await calendarService.save(calPayload);
        console.log('[SmartLink] Calendar event synced for presentation', presentation.id);
      } catch (err) {
        console.warn('[SmartLink] Could not sync calendar event:', err.message);
        // Don't block the presentation save — just log the warning
      }
    }

    /**
     * Remove a calendar event that was linked to a deleted presentation.
     */
    async function removePresentationFromCalendar(presentationId) {
      if (!calendarService) return;
      const calEventId = 'pres_' + presentationId;
      try {
        await calendarService.remove(calEventId);
        console.log('[SmartLink] Removed calendar event for presentation', presentationId);
      } catch (err) {
        // Event might not exist — that's fine
        console.warn('[SmartLink] Could not remove calendar event:', err.message);
      }
    }

    // ---- Edit (global) ----
    function editPresentation(id) {
      openPresentationModal(id);
    }

    // ---- Attach form listeners ----
    function attachFormListeners() {
      // Form submit
      if (elements.presentationForm) {
        elements.presentationForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          await onSavePresentation(new FormData(e.target));
        });
      }

      // Add button
      if (elements.addPresentationBtn) {
        elements.addPresentationBtn.addEventListener('click', () => {
          openPresentationModal();
        });
      }

      // Reset form when modal is hidden
      const modalEl = document.getElementById('presentationModal');
      if (modalEl) {
        modalEl.addEventListener('hidden.bs.modal', resetPresentationForm);
      }
    }

    attachFormListeners();
    initDetailsModal();

    // Expose edit/delete/details as global functions for onclick handlers in row HTML
    window.editPresentation = editPresentation;
    window.deletePresentation = deletePresentation;
    window.showPresentationDetails = showPresentationDetails;

    return {
      setPresentations,
      getPresentations,
      renderPresentations,
      openPresentationModal,
      onSavePresentation,
    };
  }

  window.PresentationsFeature = { init };
})(window);
