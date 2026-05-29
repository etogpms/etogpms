
(function (window) {
    // Dependencies
    let service;
    let auth;
    let utils;
    let isAdminRef;

    // State
    let currentDate = new Date(); // Tracks the month being viewed
    let events = [];
    let unsubscribeFunc = null;

    // DOM Elements
    let els = {};

    const CalendarFeature = {
        init: function ({ elements, calendarService, auth: authInstance, utils: utilsInstance, isAdminRef: isAdminRefFn }) {
            els = elements;
            service = calendarService;
            auth = authInstance;
            utils = utilsInstance;
            isAdminRef = isAdminRefFn;

            // Cache internal elements
            els.calendarSection = document.getElementById('calendarSection');
            els.calendarGrid = document.getElementById('calendarGrid');
            els.calMonthLabel = document.getElementById('calMonthLabel');

            // Modal
            const modalEl = document.getElementById('calendarModal');
            els.calModal = modalEl ? new bootstrap.Modal(modalEl) : null;
            els.calForm = document.getElementById('calendarForm');
            els.deleteBtn = document.getElementById('deleteEventBtn');
            els.saveBtn = document.getElementById('saveEventBtn');
            els.editWarn = document.getElementById('calEditWarn');

            // Bind nav controls
            const bind = (id, fn) => document.getElementById(id)?.addEventListener('click', fn);
            bind('calPrevBtn', () => changeMonth(-1));
            bind('calNextBtn', () => changeMonth(1));
            bind('calTodayBtn', () => { currentDate = new Date(); render(); });
            bind('addEventBtn', () => openModal());

            // Form Bindings
            if (els.calForm) {
                els.calForm.addEventListener('submit', onSave);
            }
            if (els.deleteBtn) {
                els.deleteBtn.addEventListener('click', onDelete);
            }

            // All Day Toggle
            const allDayCheck = document.getElementById('calAllDay');
            const timeRow = document.getElementById('calTimeRow');
            if (allDayCheck && timeRow) {
                allDayCheck.addEventListener('change', (e) => {
                    timeRow.style.display = e.target.checked ? 'none' : 'flex';
                    document.getElementById('calStartTime').required = !e.target.checked;
                    document.getElementById('calEndTime').required = !e.target.checked;
                });
            }

            // Init Day View Modal
            initDayViewModal();

            // Initial Subscribe
            subscribeToEvents();

            return this;
        },

        render: render
    };

    function changeMonth(delta) {
        currentDate.setMonth(currentDate.getMonth() + delta);
        render();
    }

    function subscribeToEvents() {
        if (unsubscribeFunc) unsubscribeFunc();

        // Subscribe to ALL events for now to support multi-month ranges easily
        // In production with large data, query by range
        unsubscribeFunc = service.subscribe((snap) => {
            events = snap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
            render();
        });
    }

    function render() {
        if (!els.calendarGrid) return;

        // Header
        const year = currentDate.getFullYear();
        const month = currentDate.getMonth();
        const monthName = currentDate.toLocaleString('default', { month: 'long' });
        if (els.calMonthLabel) els.calMonthLabel.textContent = `${monthName} ${year}`;

        // Grid calc
        const firstDay = new Date(year, month, 1);
        const lastDay = new Date(year, month + 1, 0);
        const startDayOfWeek = firstDay.getDay(); // 0=Sun
        const daysInMonth = lastDay.getDate();

        let html = `
      <div class="cal-header">Sun</div>
      <div class="cal-header">Mon</div>
      <div class="cal-header">Tue</div>
      <div class="cal-header">Wed</div>
      <div class="cal-header">Thu</div>
      <div class="cal-header">Fri</div>
      <div class="cal-header">Sat</div>
    `;

        // Padding before first day
        for (let i = 0; i < startDayOfWeek; i++) {
            html += `<div class="cal-grid-cell bg-light border-end border-bottom"></div>`;
        }

        // Today check
        const today = new Date();
        const isTodayYear = today.getFullYear() === year;
        const isTodayMonth = today.getMonth() === month;
        const todayDate = today.getDate();

        // Days extraction
        for (let d = 1; d <= daysInMonth; d++) {
            const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const isToday = isTodayYear && isTodayMonth && (d === todayDate);

            // Filter events: Date is between startDate and endDate
            const dayEvents = events.filter(ev => {
                return dateStr >= ev.startDate && dateStr <= ev.endDate;
            }).sort((a, b) => {
                // Sort by time (all day first)
                if (a.allDay && !b.allDay) return -1;
                if (!a.allDay && b.allDay) return 1;
                return (a.startTime || '').localeCompare(b.startTime || '');
            });

            let eventsHtml = '';
            dayEvents.forEach(ev => {
                const typeClass = `type-${ev.type || 'activity'}`;
                let timeLabel = '';
                if (!ev.allDay && ev.startTime) {
                    timeLabel = formatTime12(ev.startTime);
                }

                // Escape content
                const safeTitle = (ev.title || '').replace(/</g, '&lt;');

                eventsHtml += `<div class="cal-event ${typeClass}" onclick="window.Calendar.openEvent('${ev.id}', event)" title="${safeTitle} \n${timeLabel}">
          ${timeLabel ? `<span style="font-size:0.7em; opacity:0.8;">${timeLabel}</span> ` : ''}
          ${safeTitle}
        </div>`;
            });

            html += `<div class="cal-grid-cell cal-day ${isToday ? 'today' : ''}" onclick="window.Calendar.openDate('${dateStr}')">
        <div class="cal-date-num">${d}</div>
        <div class="cal-events-list">
          ${eventsHtml}
        </div>
      </div>`;
        }

        els.calendarGrid.innerHTML = html;
    }

    // --- Modal & CRUD ---

    function openModal(eventData = null, defaultDate = null) {
        const form = els.calForm;
        form.reset();
        const isEdit = !!eventData;
        const currentUser = auth.currentUser;

        // Check permission
        const isAdmin = isAdminRef ? isAdminRef() : false; // Value from script.js
        const isOwner = isEdit && currentUser && (eventData.ownerUid === currentUser.uid);

        // Explicit email check as fallback/ensure
        const adminEmail = 'johnlowel.fradejas@mwss.gov.ph';
        const isUserAdminEmail = currentUser && currentUser.email === adminEmail;

        const canEdit = !isEdit || isOwner || isAdmin || isUserAdminEmail;

        document.getElementById('calEventId').value = isEdit ? eventData.id : '';
        document.getElementById('calTitle').value = isEdit ? eventData.title : '';
        document.getElementById('calType').value = isEdit ? eventData.type : 'meeting';
        document.getElementById('calDesc').value = isEdit ? (eventData.description || '') : '';
        document.getElementById('calDocLink').value = isEdit ? (eventData.docLink || '') : '';

        // Dates
        const todayStr = new Date().toISOString().split('T')[0];
        document.getElementById('calStartDate').value = isEdit ? eventData.startDate : (defaultDate || todayStr);
        document.getElementById('calEndDate').value = isEdit ? eventData.endDate : (defaultDate || todayStr);

        // Time & All Day
        const allDayCheck = document.getElementById('calAllDay');
        allDayCheck.checked = isEdit ? !!eventData.allDay : false;

        document.getElementById('calStartTime').value = isEdit ? (eventData.startTime || '') : '';
        document.getElementById('calEndTime').value = isEdit ? (eventData.endTime || '') : '';

        // Toggle time inputs
        allDayCheck.dispatchEvent(new Event('change'));

        // Buttons Visibility
        els.deleteBtn.style.display = isEdit && canEdit ? 'block' : 'none';
        els.saveBtn.style.display = canEdit ? 'block' : 'none';
        els.editWarn.classList.toggle('d-none', canEdit);

        // Disable inputs if view-only
        const inputs = form.querySelectorAll('input, select, textarea');
        inputs.forEach(el => el.disabled = !canEdit);

        els.calModal.show();
    }

    async function onSave(e) {
        e.preventDefault();
        if (!auth.currentUser) return alert('Must be logged in.');

        const formData = new FormData(e.target);

        // Validation: End Date >= Start Date
        if (formData.get('calEndDate') < formData.get('calStartDate')) {
            return alert('End Date cannot be before Start Date');
        }

        const eventId = formData.get('calEventId');
        const isEdit = !!eventId;

        // Build Object
        const data = {
            title: formData.get('calTitle'),
            type: formData.get('calType'),
            startDate: formData.get('calStartDate'),
            endDate: formData.get('calEndDate'),
            allDay: formData.get('calAllDay') === 'on',
            startTime: formData.get('calStartTime') || null,
            endTime: formData.get('calEndTime') || null,
            description: formData.get('calDesc'),
            docLink: formData.get('calDocLink') || null,
            updatedAt: window.AppFirebase.serverTimestamp()
        };

        if (!isEdit) {
            data.ownerUid = auth.currentUser.uid;
            data.ownerName = auth.currentUser.displayName || auth.currentUser.email;
            data.createdAt = window.AppFirebase.serverTimestamp();
            data.id = `evt_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`;
        } else {
            data.id = eventId;
            // Don't overwrite ownerUid
        }

        try {
            const ref = service.collection().doc(data.id);
            if (isEdit) {
                await ref.update(data);
            } else {
                await ref.set(data);
            }
            els.calModal.hide();
            if (window.showToast) window.showToast('Event saved successfully!', 'success');

            // Trigger immediate notification check for this event
            try {
                if (window.NotificationsFeature && window.NotificationsFeature.startListening && auth.currentUser) {
                    // Re-trigger notification generation to pick up newly saved event
                    window.NotificationsFeature.startListening(auth.currentUser);
                }
            } catch (_) { /* noop */ }
        } catch (err) {
            console.error(err);
            alert('Error saving event: ' + err.message);
        }
    }

    async function onDelete() {
        const id = document.getElementById('calEventId').value;
        if (!id) return;
        if (!confirm('Are you sure you want to delete this event?')) return;

        try {
            await service.remove(id);
            els.calModal.hide();
        } catch (err) {
            console.error(err);
            alert('Delete failed: ' + err.message);
        }
    }

    function formatTime12(timeStr) {
        if (!timeStr) return '';
        const [h, m] = timeStr.split(':');
        const hour = parseInt(h, 10);
        const ampm = hour >= 12 ? 'PM' : 'AM';
        const h12 = hour % 12 || 12;
        return `${h12}:${m} ${ampm}`;
    }

    // --- Day View Modal ---

    let dayModal = null;
    let currentDayViewDate = null;

    function initDayViewModal() {
        const dayModalEl = document.getElementById('calendarDayModal');
        if (dayModalEl) {
            dayModal = new bootstrap.Modal(dayModalEl);
        }

        // "Add Event" button inside day view
        const addBtn = document.getElementById('dayViewAddBtn');
        if (addBtn) {
            addBtn.addEventListener('click', () => {
                if (dayModal) dayModal.hide();
                setTimeout(() => openModal(null, currentDayViewDate), 300);
            });
        }
    }

    function openDayView(dateStr) {
        currentDayViewDate = dateStr;

        // Format the date nicely for the header
        const [y, mon, day] = dateStr.split('-').map(Number);
        const dateObj = new Date(y, mon - 1, day);
        const formatted = dateObj.toLocaleDateString('en-US', {
            weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
        });
        const label = document.getElementById('dayViewDateLabel');
        if (label) label.textContent = formatted;

        // Filter events for this date
        const dayEvents = events.filter(ev => {
            return dateStr >= ev.startDate && dateStr <= ev.endDate;
        }).sort((a, b) => {
            if (a.allDay && !b.allDay) return -1;
            if (!a.allDay && b.allDay) return 1;
            return (a.startTime || '').localeCompare(b.startTime || '');
        });

        const listEl = document.getElementById('dayViewEventsList');
        const emptyEl = document.getElementById('dayViewEmpty');
        const currentUser = auth.currentUser;
        const isAdmin = isAdminRef ? isAdminRef() : false;
        const adminEmail = 'johnlowel.fradejas@mwss.gov.ph';
        const isUserAdminEmail = currentUser && currentUser.email === adminEmail;

        if (dayEvents.length === 0) {
            if (listEl) listEl.innerHTML = '';
            if (emptyEl) emptyEl.style.display = 'block';
        } else {
            if (emptyEl) emptyEl.style.display = 'none';

            let html = '';
            dayEvents.forEach(ev => {
                const safeTitle = (ev.title || '').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                const safeDesc = (ev.description || '').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                const typeName = (ev.type || 'activity').charAt(0).toUpperCase() + (ev.type || 'activity').slice(1);
                const typeClass = `type-${ev.type || 'activity'}`;

                // Time display
                let timeMeta = '';
                if (ev.allDay) {
                    timeMeta = '<i class="fa-regular fa-clock me-1"></i>All Day';
                } else if (ev.startTime) {
                    timeMeta = `<i class="fa-regular fa-clock me-1"></i>${formatTime12(ev.startTime)}`;
                    if (ev.endTime) timeMeta += ` – ${formatTime12(ev.endTime)}`;
                }

                // Owner check for edit/delete buttons
                const isOwner = currentUser && (ev.ownerUid === currentUser.uid);
                const canManage = isOwner || isAdmin || isUserAdminEmail;

                let actionsHtml = '';
                if (canManage) {
                    actionsHtml = `
                        <button class="btn btn-sm btn-outline-primary py-0 px-2" onclick="window.Calendar.editEvent('${ev.id}')" title="Edit">
                            <i class="fa fa-pencil"></i>
                        </button>
                        <button class="btn btn-sm btn-outline-danger py-0 px-2" onclick="window.Calendar.deleteEvent('${ev.id}')" title="Delete">
                            <i class="fa fa-trash"></i>
                        </button>`;
                }

                html += `
                <div class="day-view-item">
                    <div class="day-view-type-dot ${typeClass}"></div>
                    <div class="day-view-info">
                        <div class="day-view-title">${safeTitle}</div>
                        <div class="day-view-meta">
                            <span class="badge bg-${ev.type === 'meeting' ? 'primary' : ev.type === 'leave' ? 'warning text-dark' : 'success'} me-1" style="font-size:0.7rem;">${typeName}</span>
                            ${timeMeta}
                        </div>
                        ${safeDesc ? `<div class="day-view-desc">${safeDesc}</div>` : ''}
                        ${ev.docLink ? `<div class="day-view-doc-link"><i class="fa-solid fa-link me-1"></i><a href="${ev.docLink}" target="_blank" rel="noopener noreferrer">View Document</a></div>` : ''}
                        <div class="day-view-owner"><i class="fa-regular fa-user me-1"></i>${ev.ownerName || 'Unknown'}</div>
                    </div>
                    <div class="day-view-actions">
                        ${actionsHtml}
                    </div>
                </div>`;
            });

            if (listEl) listEl.innerHTML = html;
        }

        // Show "Add Event" button only for signed-in users
        const addBtn = document.getElementById('dayViewAddBtn');
        if (addBtn) addBtn.style.display = currentUser ? 'inline-block' : 'none';

        if (dayModal) dayModal.show();
    }

    // Expose Globals
    window.Calendar = {
        openEvent: (id, e) => {
            if (e) e.stopPropagation();
            const ev = events.find(x => x.id === id);
            if (ev) openModal(ev);
        },
        openDate: (dateStr) => {
            openDayView(dateStr);
        },
        editEvent: (id) => {
            const ev = events.find(x => x.id === id);
            if (!ev) return;
            if (dayModal) dayModal.hide();
            setTimeout(() => openModal(ev), 300);
        },
        deleteEvent: async (id) => {
            if (!confirm('Are you sure you want to delete this event?')) return;
            try {
                await service.remove(id);
                // Refresh the day view if still the same date
                if (currentDayViewDate) {
                    setTimeout(() => openDayView(currentDayViewDate), 300);
                }
            } catch (err) {
                console.error(err);
                alert('Delete failed: ' + err.message);
            }
        }
    };

    window.CalendarFeature = CalendarFeature;

})(window);
