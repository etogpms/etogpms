// OPCR feature module: handles CRUD, rendering, and search for OPCR entries.
// Depends on OPCRService, Firebase auth, and DOM elements passed in.
(function (window) {
    if (window.OpcrFeature) return;

    function init(opts) {
        const {
            elements,
            opcrService,
            auth,
            utils,
            isAdminRef = () => false,
        } = opts || {};

        if (!elements || !opcrService || !auth || !utils) {
            throw new Error('OpcrFeature.init missing required dependencies');
        }

        const { serverTimestamp } = window.AppFirebase || {};

        let opcrEntries = [];

        function setOpcrEntries(list) {
            opcrEntries = Array.isArray(list) ? list.slice() : [];
        }

        function getOpcrEntries() {
            return opcrEntries.slice();
        }

        // ---- Rendering ----

        function opcrRowHtml(entry) {
            const canEdit = !!auth.currentUser;
            const isAdmin = isAdminRef();

            let actionsHtml = '';
            if (canEdit) {
                actionsHtml += `<button class="btn btn-sm btn-outline-primary me-1" title="Edit Form" onclick="event.stopPropagation(); editOpcr('${entry.id}')"><i class="fa fa-pencil"></i></button>`;
            }
            if (isAdmin) {
                actionsHtml += `<button class="btn btn-sm btn-outline-danger" title="Delete" onclick="event.stopPropagation(); deleteOpcr('${entry.id}')"><i class="fa fa-trash"></i></button>`;
            }

            // Format timestamp if available
            let updatedStr = '';
            const fmtUi = window.AppUtils?.formatDateUI || (x => x);
            const nYmd = window.AppUtils?.normalizeYmd || (x => x);
            if (entry.updatedAt) {
                if (entry.updatedAt.toDate) updatedStr = fmtUi(nYmd(entry.updatedAt.toDate()));
                else if (entry.updatedAt.seconds) updatedStr = fmtUi(nYmd(new Date(entry.updatedAt.seconds * 1000)));
                else updatedStr = fmtUi(nYmd(new Date(entry.updatedAt)));
            } else if (entry.createdAt) {
                if (entry.createdAt.toDate) updatedStr = fmtUi(nYmd(entry.createdAt.toDate()));
                else if (entry.createdAt.seconds) updatedStr = fmtUi(nYmd(new Date(entry.createdAt.seconds * 1000)));
                else updatedStr = fmtUi(nYmd(new Date(entry.createdAt)));
            }

            // Status badge
            const currentStatus = entry.status || 'Draft';
            const statusClass = currentStatus === 'Approved' ? 'bg-success' :
                currentStatus === 'Submitted' ? 'bg-primary' : 'bg-secondary';

            return `<tr data-id="${entry.id}" style="cursor:pointer;" onclick="showOpcrDetailsModal('${entry.id}')" title="Click to view details">
        <td>
          <div class="fw-semibold">${entry.department || '—'}</div>
          ${entry.departmentOther ? `<div class="small text-muted">${entry.departmentOther}</div>` : ''}
        </td>
        <td>
          <div>${updatedStr}</div>
          <span class="badge ${statusClass}">${currentStatus}</span>
        </td>
        <td>${entry.year || '—'}</td>
        <td onclick="event.stopPropagation();">${actionsHtml}</td>
      </tr>`;
        }

        // Show OPCR details modal
        window.showOpcrDetailsModal = function (entryId) {
            const entry = opcrEntries.find(e => e.id === entryId);
            if (!entry) return;

            // Build modal HTML
            const modal = document.getElementById('opcrDetailsModal');
            if (!modal) {
                console.error('OPCR Details Modal not found');
                return;
            }

            // Department title
            document.getElementById('opcrModalDepartment').textContent = entry.department || entry.departmentOther || 'Unknown Department';
            document.getElementById('opcrModalYear').textContent = entry.year || '';

            // Status dropdown
            const currentStatus = entry.status || 'Draft';
            const statusSelect = document.getElementById('opcrModalStatus');
            if (statusSelect) {
                statusSelect.value = currentStatus;
                statusSelect.dataset.entryId = entryId;
            }

            // Build edit logs with restore buttons for versions
            const logsContainer = document.getElementById('opcrModalLogs');
            if (logsContainer) {
                // Combine versions and legacy editLogs for display
                const versions = entry.versions || [];
                const legacyLogs = entry.editLogs || [];

                // Build combined history (newest first)
                let historyHtml = '';

                // First show version entries with restore buttons (newest first)
                if (versions.length > 0) {
                    const sortedVersions = [...versions].sort((a, b) =>
                        new Date(b.timestamp) - new Date(a.timestamp)
                    );

                    sortedVersions.forEach((ver, idx) => {
                        let verTime = '';
                        if (ver.timestamp) {
                            verTime = new Date(ver.timestamp).toLocaleString();
                        }

                        // Only show restore button if not the most recent version
                        const isLatest = idx === 0;
                        const restoreBtn = !isLatest ? `
                            <button class="btn btn-sm btn-outline-success ms-2" 
                                onclick="event.stopPropagation(); restoreOpcrVersion('${entry.id}', '${ver.versionId}')"
                                title="Restore to this version">
                                <i class="fa fa-undo me-1"></i>Restore
                            </button>
                        ` : `<span class="badge bg-info ms-2">Current</span>`;

                        historyHtml += `
                            <div class="d-flex justify-content-between align-items-start py-2 border-bottom">
                                <div class="flex-grow-1">
                                    <strong>${ver.action || 'Form saved'}</strong>
                                    <div class="small text-muted">by ${ver.user || 'Unknown'}</div>
                                </div>
                                <div class="text-end d-flex align-items-center">
                                    <span class="small text-muted">${verTime}</span>
                                    ${restoreBtn}
                                </div>
                            </div>
                        `;
                    });
                }

                // Then show legacy logs without restore (older entries without snapshots)
                const legacyOnlyLogs = legacyLogs.filter(log => {
                    // Don't show if there's a matching version
                    return !versions.some(v => v.timestamp === log.timestamp);
                });

                if (legacyOnlyLogs.length > 0) {
                    const sortedLogs = [...legacyOnlyLogs].sort((a, b) =>
                        new Date(b.timestamp) - new Date(a.timestamp)
                    );

                    sortedLogs.forEach(log => {
                        let logTime = '';
                        if (log.timestamp) {
                            if (log.timestamp.toDate) logTime = log.timestamp.toDate().toLocaleString();
                            else if (log.timestamp.seconds) logTime = new Date(log.timestamp.seconds * 1000).toLocaleString();
                            else logTime = new Date(log.timestamp).toLocaleString();
                        }

                        historyHtml += `
                            <div class="d-flex justify-content-between align-items-start py-2 border-bottom">
                                <div class="flex-grow-1">
                                    <strong>${log.action || 'Edit'}</strong>
                                    <div class="small text-muted">by ${log.user || 'Unknown'}</div>
                                </div>
                                <div class="text-end d-flex align-items-center">
                                    <span class="small text-muted">${logTime}</span>
                                    <span class="badge bg-secondary ms-2" title="No snapshot available">No restore</span>
                                </div>
                            </div>
                        `;
                    });
                }

                if (historyHtml) {
                    logsContainer.innerHTML = historyHtml;
                } else {
                    logsContainer.innerHTML = '<div class="text-muted text-center py-3"><i class="fa fa-info-circle me-1"></i>No edit history available</div>';
                }
            }

            // Show modal (Bootstrap)
            const bsModal = new bootstrap.Modal(modal);
            bsModal.show();
        };

        // Restore OPCR to a specific version
        window.restoreOpcrVersion = async function (entryId, versionId) {
            const entry = opcrEntries.find(e => e.id === entryId);
            if (!entry) {
                alert('Entry not found.');
                return;
            }

            const versions = entry.versions || [];
            const version = versions.find(v => v.versionId === versionId);
            if (!version || !version.snapshot) {
                alert('Version snapshot not found. Cannot restore.');
                return;
            }

            // Confirm restore
            const versionTime = new Date(version.timestamp).toLocaleString();
            if (!confirm(`Are you sure you want to restore to the version from ${versionTime}?\n\nThis will replace the current form data with the saved snapshot.`)) {
                return;
            }

            try {
                const user = auth.currentUser;
                if (!user) {
                    alert('You must be logged in to restore versions.');
                    return;
                }

                // Get current data for new version entry
                const currentLogs = entry.editLogs || [];
                const currentVersions = entry.versions || [];

                // Create a restore log entry
                const restoreLog = {
                    action: `Restored to version from ${versionTime}`,
                    user: user.email,
                    timestamp: new Date().toISOString()
                };

                // Create a new version entry for the restore action
                const restoreVersion = {
                    versionId: 'v' + Date.now(),
                    timestamp: new Date().toISOString(),
                    user: user.email,
                    action: `Restored to version from ${versionTime}`,
                    snapshot: version.snapshot // Keep the same snapshot data
                };

                // Prepare update data with restored snapshot fields
                const updateData = {
                    department: version.snapshot.department,
                    year: version.snapshot.year,
                    periodStart: version.snapshot.periodStart,
                    periodEnd: version.snapshot.periodEnd,
                    items: version.snapshot.items,
                    signatoriesPartI: version.snapshot.signatoriesPartI,
                    signatories: version.snapshot.signatories,
                    editLogs: [...currentLogs, restoreLog],
                    versions: [...currentVersions, restoreVersion].slice(-5),
                    updatedAt: serverTimestamp(),
                    updatedBy: user.email
                };

                await opcrService.update(entryId, updateData);

                // Close modal
                const modal = bootstrap.Modal.getInstance(document.getElementById('opcrDetailsModal'));
                if (modal) modal.hide();

                alert('Successfully restored to the selected version. Refresh the form to see changes.');

            } catch (err) {
                console.error('Error restoring OPCR version:', err);
                alert('Failed to restore version: ' + err.message);
            }
        };

        // Save status from modal
        window.saveOpcrStatusFromModal = async function () {
            const statusSelect = document.getElementById('opcrModalStatus');
            if (!statusSelect) return;

            const entryId = statusSelect.dataset.entryId;
            const newStatus = statusSelect.value;

            try {
                const user = auth.currentUser;
                if (!user) {
                    alert('You must be logged in to change status.');
                    return;
                }

                // Get current entry to append to edit logs
                const entry = opcrEntries.find(e => e.id === entryId);
                const currentLogs = (entry && entry.editLogs) || [];

                const updateData = {
                    status: newStatus,
                    updatedAt: serverTimestamp(),
                    updatedBy: user.email,
                    editLogs: [...currentLogs, {
                        action: `Status changed to ${newStatus}`,
                        user: user.email,
                        timestamp: new Date().toISOString()
                    }]
                };

                await opcrService.update(entryId, updateData);

                // Close modal
                const modal = bootstrap.Modal.getInstance(document.getElementById('opcrDetailsModal'));
                if (modal) modal.hide();

                console.log(`OPCR ${entryId} status updated to ${newStatus}`);
            } catch (err) {
                console.error('Error updating OPCR status:', err);
                alert('Failed to update status: ' + err.message);
            }
        };

        function renderOpcr() {
            const tbody = elements.opcrTbody;
            const empty = elements.opcrEmpty;
            if (!tbody) return;

            let list = opcrEntries.slice();

            // Filter by Search
            const search = (elements.opcrSearchInput?.value || '').toLowerCase();
            if (search) {
                list = list.filter(e =>
                    (e.department || '').toLowerCase().includes(search) ||
                    (e.departmentOther || '').toLowerCase().includes(search) ||
                    (e.status || '').toLowerCase().includes(search) ||
                    (e.remarks || '').toLowerCase().includes(search)
                );
            }

            // Filter by Year
            const yearFilter = elements.opcrYearFilter?.value;
            if (yearFilter && yearFilter !== '') {
                list = list.filter(e => String(e.year) === String(yearFilter));
            }

            // Sort by Year desc, then Department asc
            list.sort((a, b) => {
                if (b.year !== a.year) return b.year - a.year;
                return (a.department || '').localeCompare(b.department || '');
            });

            if (list.length === 0) {
                tbody.innerHTML = '';
                if (empty) empty.style.display = 'block';
            } else {
                if (empty) empty.style.display = 'none';
                tbody.innerHTML = list.map(opcrRowHtml).join('');
            }

            // Update year filter options if needed (dynamic population)
            populateOpcrYearOptions();
        }

        function populateOpcrYearOptions() {
            const yearSelect = elements.opcrYearFilter;
            const formYearSelect = elements.opcrYear;
            if (!yearSelect) return;

            // Collect years from entries + current year + next year
            const years = new Set();
            const currentYear = new Date().getFullYear();
            years.add(currentYear);
            years.add(currentYear + 1);
            opcrEntries.forEach(e => { if (e.year) years.add(Number(e.year)); });

            const sortedYears = Array.from(years).sort((a, b) => b - a);

            // Preserve current selection
            const currentVal = yearSelect.value;

            // Rebuild filter options
            // Check if we need to rebuild? Simple check: if length differs or first item differs
            // Ideally we shouldn't rebuild constantly to avoid losing selection or flickering, 
            // but simplistic approach for now:
            // Only rebuild if option count changed significantly or empty
            if (yearSelect.options.length <= 1) { // only "All" exists
                yearSelect.innerHTML = '<option value="">All</option>' +
                    sortedYears.map(y => `<option value="${y}">${y}</option>`).join('');
                yearSelect.value = currentVal;
            }

            // Also populate the form year select
            if (formYearSelect && formYearSelect.options.length === 0) {
                formYearSelect.innerHTML = sortedYears.map(y => `<option value="${y}">${y}</option>`).join('');
                formYearSelect.value = currentYear;
            }
        }

        // ---- CRUD ----

        let currentOpcrId = null;

        // ---- Standalone Form Logic ----

        function switchToOpcrList() {
            document.getElementById('opcrSection').style.display = 'block';
            document.getElementById('opcrFormSection').style.display = 'none';
            currentOpcrId = null;
        }

        function showOpcrForm(entry) {
            document.getElementById('opcrSection').style.display = 'none';
            document.getElementById('opcrFormSection').style.display = 'block';

            // Set current entry ID
            currentOpcrId = entry ? entry.id : null;

            // Populate form metadata
            const deptInput = document.getElementById('opcrFormDepartment');
            const periodStartSel = document.getElementById('opcrFormPeriodStart');
            const periodEndSel = document.getElementById('opcrFormPeriodEnd');
            const yearInput = document.getElementById('opcrFormYear');
            const rateeNameInput = document.getElementById('opcrRateeName');
            const rateePosInput = document.getElementById('opcrRateePosition');
            const footerRatee = document.getElementById('opcrFooterRatee');
            const footerRateePos = document.getElementById('opcrFooterRateePos');

            if (entry) {
                if (deptInput) deptInput.value = entry.department || entry.departmentOther || '';
                if (periodStartSel) periodStartSel.value = entry.periodStart || 'JANUARY';
                if (periodEndSel) periodEndSel.value = entry.periodEnd || 'DECEMBER';
                if (yearInput) yearInput.value = entry.year || new Date().getFullYear();
                if (rateeNameInput) rateeNameInput.value = entry.rateeName || '';
                if (rateePosInput) rateePosInput.value = entry.rateePosition || '';
                if (footerRatee) footerRatee.textContent = entry.rateeName || '-';
                if (footerRateePos) footerRateePos.textContent = entry.rateePosition || 'Position';
            } else {
                // Clear fields for new entry
                if (deptInput) deptInput.value = '';
                if (yearInput) yearInput.value = new Date().getFullYear();
                if (rateeNameInput) rateeNameInput.value = '';
                if (rateePosInput) rateePosInput.value = '';
            }

            // Populate table rows
            const tbody = document.getElementById('opcrFormTbody');
            tbody.innerHTML = '';

            if (entry && entry.items && Array.isArray(entry.items) && entry.items.length > 0) {
                entry.items.forEach(item => {
                    if (item.type === 'section') {
                        addOpcrSectionHeader(item.title || '');
                    } else {
                        addOpcrRow(item);
                    }
                });
            } else {
                // Default: add one section header and one empty row
                addOpcrSectionHeader('A. Major Final Outputs (MFOs)/Operations');
                addOpcrRow();
            }

            // Update status badge
            const badge = document.getElementById('formOpcrStatusBadge');
            if (badge) {
                badge.textContent = entry ? (entry.status || 'Draft') : 'Draft';
                badge.className = `badge ${entry && entry.status === 'Submitted' ? 'bg-success' : 'bg-warning text-dark'} me-2`;
            }
        }

        function addOpcrRow(data = {}) {
            const tbody = document.getElementById('opcrFormTbody');
            const rowId = 'row-' + Date.now() + Math.random().toString(36).substr(2, 5);

            const tr = document.createElement('tr');
            tr.id = rowId;
            tr.className = 'opcr-item-row';
            tr.style.pageBreakInside = 'avoid';

            // Helper for rating calculation
            const calcAve = `
            const row = this.closest('tr');
            const q = parseFloat(row.querySelector('.rating-q').innerText || 0);
            const e = parseFloat(row.querySelector('.rating-e').innerText || 0);
            const t = parseFloat(row.querySelector('.rating-t').innerText || 0);
            const count = (q>0?1:0) + (e>0?1:0) + (t>0?1:0);
            const ave = count > 0 ? ((q+e+t)/count).toFixed(2) : '';
            row.querySelector('.rating-ave').innerText = ave;
        `;
            // Updated structure to match Excel template
            tr.innerHTML = `
            <td class="p-1">
                <textarea class="form-control form-control-sm border-0 rounded-0" rows="3" placeholder="MFO / Performance Indicator">${data.mfo || ''}</textarea>
            </td>
            <td class="p-1">
                <input type="text" class="form-control form-control-sm border-0 text-end" placeholder="Budget" value="${data.budget || ''}">
            </td>
            <td class="p-1">
                <textarea class="form-control form-control-sm border-0 rounded-0" rows="3" placeholder="Success Indicator / Targets">${data.indicators || ''}</textarea>
            </td>
            <td class="p-1">
                <input type="text" class="form-control form-control-sm border-0 text-center" placeholder="Division/Person" value="${data.responsible || ''}">
            </td>
            <td contenteditable="true" class="rating-q text-center align-middle" oninput="${calcAve}">${data.ratingQ || ''}</td>
            <td contenteditable="true" class="rating-e text-center align-middle" oninput="${calcAve}">${data.ratingE || ''}</td>
            <td contenteditable="true" class="rating-t text-center align-middle" oninput="${calcAve}">${data.ratingT || ''}</td>
            <td class="rating-ave text-center align-middle bg-light fw-bold">${data.ratingAve || ''}</td>
            <td class="p-1">
                <textarea class="form-control form-control-sm border-0" rows="3" placeholder="Accomplishments / Updates">${data.accomplishments || ''}</textarea>
            </td>
            <td class="text-center align-middle">
                <button class="btn btn-sm text-danger opcr-action-btn" onclick="this.closest('tr').remove()" title="Remove Row">
                    <i class="fa fa-times"></i>
                </button>
            </td>
        `;

            tbody.appendChild(tr);
        }

        // Add section header row (e.g., "A. Major Final Outputs (MFOs)")
        function addOpcrSectionHeader(title = '') {
            const tbody = document.getElementById('opcrFormTbody');
            const rowId = 'section-' + Date.now();

            const tr = document.createElement('tr');
            tr.id = rowId;
            tr.className = 'opcr-section-header table-secondary';
            tr.innerHTML = `
            <td colspan="9" class="p-2">
                <div class="d-flex align-items-center">
                    <input type="text" class="form-control form-control-sm border-0 bg-transparent fw-bold" 
                        placeholder="Section Header (e.g., A. Major Final Outputs (MFOs)/Operations)" 
                        value="${title}">
                    <button class="btn btn-sm text-danger ms-2" onclick="this.closest('tr').remove()" title="Remove Section">
                        <i class="fa fa-times"></i>
                    </button>
                </div>
            </td>
        `;
            tbody.appendChild(tr);
        }

        // Expose globally
        window.addOpcrSectionHeader = addOpcrSectionHeader;

        async function saveOpcrForm() {
            if (!currentOpcrId) {
                alert('Error: No OPCR Entry ID. Please create an entry from the dashboard first.');
                return;
            }

            const user = auth.currentUser;
            if (!user) { alert('You must be logged in.'); return; }

            // Gather Form Metadata
            const formDept = document.getElementById('opcrFormDepartment')?.value || '';
            const formPeriodStart = document.getElementById('opcrFormPeriodStart')?.value || 'JANUARY';
            const formPeriodEnd = document.getElementById('opcrFormPeriodEnd')?.value || 'DECEMBER';
            const formYear = document.getElementById('opcrFormYear')?.value || new Date().getFullYear();
            const rateeName = document.getElementById('opcrRateeName')?.value || '';
            const rateePosition = document.getElementById('opcrRateePosition')?.value || '';

            // Gather Table Data (rows and section headers)
            const allRows = document.querySelectorAll('#opcrFormTbody tr');
            const items = [];

            allRows.forEach(row => {
                if (row.classList.contains('opcr-section-header')) {
                    // Section header
                    const headerInput = row.querySelector('input');
                    items.push({
                        type: 'section',
                        title: headerInput ? headerInput.value : ''
                    });
                } else if (row.classList.contains('opcr-item-row')) {
                    // Data row - new structure with Budget column
                    const mfo = row.cells[0]?.querySelector('textarea')?.value || '';
                    const budget = row.cells[1]?.querySelector('input')?.value || '';
                    const indicators = row.cells[2]?.querySelector('textarea')?.value || '';
                    const responsible = row.cells[3]?.querySelector('input')?.value || '';
                    const ratingQ = row.querySelector('.rating-q')?.innerText?.trim() || '';
                    const ratingE = row.querySelector('.rating-e')?.innerText?.trim() || '';
                    const ratingT = row.querySelector('.rating-t')?.innerText?.trim() || '';
                    const ratingAve = row.querySelector('.rating-ave')?.innerText?.trim() || '';
                    const accomplishments = row.cells[8]?.querySelector('textarea')?.value || '';

                    items.push({
                        type: 'item',
                        mfo, budget, indicators, responsible,
                        ratingQ, ratingE, ratingT, ratingAve,
                        accomplishments
                    });
                }
            });

            // Update the existing doc
            try {
                await opcrService.collection().doc(currentOpcrId).update({
                    department: formDept,
                    periodStart: formPeriodStart,
                    periodEnd: formPeriodEnd,
                    year: Number(formYear),
                    rateeName: rateeName,
                    rateePosition: rateePosition,
                    items: items,
                    updatedAt: serverTimestamp ? serverTimestamp() : new Date(),
                    updatedBy: user.email
                });
                alert('OPCR Form saved successfully.');
            } catch (err) {
                console.error(err);
                alert('Failed to save OPCR form: ' + err.message);
            }
        }

        // Override editOpcr to open standalone form in new tab
        function editOpcr(id) {
            const entry = opcrEntries.find(e => e.id === id);
            if (!entry) return;

            // Open standalone OPCR form in new tab (uses same origin, shares auth state)
            const formUrl = `opcr-form.html?id=${id}`;
            window.open(formUrl, '_blank');
        }

        function resetOpcrForm() {
            // Existing reset logic for modal if still used for creation
            if (elements.opcrForm) elements.opcrForm.reset();
            if (elements.opcrId) elements.opcrId.value = '';
            if (elements.opcrDepartmentOther) elements.opcrDepartmentOther.style.display = 'none';
            if (elements.opcrYear) elements.opcrYear.value = new Date().getFullYear();
            // Reset group dropdown
            const groupSelect = document.getElementById('opcrGroup');
            if (groupSelect) groupSelect.selectedIndex = 0;
            // Clear signatory fields
            const deptManager = document.getElementById('opcrDeptManager');
            const deputyAdmin = document.getElementById('opcrDeputyAdmin');
            const pmtChair = document.getElementById('opcrPmtChair');
            const admin = document.getElementById('opcrAdmin');
            if (deptManager) deptManager.value = '';
            if (deputyAdmin) deputyAdmin.value = '';
            if (pmtChair) pmtChair.value = '';
            if (admin) admin.value = '';
        }

        async function deleteOpcr(id) {
            if (!isAdminRef()) {
                alert('Only administrators can delete OPCR entries.');
                return;
            }
            if (!confirm('Are you sure you want to delete this OPCR entry?')) return;
            try {
                await opcrService.remove(id);
                // UI update handled by subscription
            } catch (err) {
                console.error(err);
                alert('Failed to delete OPCR entry: ' + err.message);
            }
        }

        async function handleOpcrSubmit(e) {
            e.preventDefault();
            const fd = new FormData(elements.opcrForm);
            const id = fd.get('opcrId');
            const department = fd.get('opcrDepartment') || (elements.opcrDepartment ? elements.opcrDepartment.value : '');
            const group = document.getElementById('opcrGroup')?.value || '';
            let departmentOther = '';

            if (department === 'Others') {
                departmentOther = (document.getElementById('opcrDepartmentOther')?.value || '').trim();
                if (!departmentOther) {
                    alert('Please specify the department.');
                    return;
                }
            }

            const year = Number(document.getElementById('opcrYear')?.value);
            const status = document.getElementById('opcrStatus')?.value;
            const remarks = document.getElementById('opcrRemarks')?.value;

            // Signatory fields
            const deptManager = document.getElementById('opcrDeptManager')?.value || '';
            const deputyAdmin = document.getElementById('opcrDeputyAdmin')?.value || '';
            const pmtChair = document.getElementById('opcrPmtChair')?.value || '';
            const administrator = document.getElementById('opcrAdmin')?.value || '';

            const user = auth.currentUser;
            if (!user) { alert('You must be logged in.'); return; }

            const payload = {
                group,
                department: department,
                departmentOther: (department === 'Others') ? departmentOther : null,
                year,
                status,
                remarks,
                // Signatory fields
                deptManager,
                deputyAdmin,
                pmtChair,
                administrator,
                updatedAt: serverTimestamp ? serverTimestamp() : new Date(),
                updatedBy: user.email
            };

            if (!id) {
                payload.createdAt = serverTimestamp ? serverTimestamp() : new Date();
                payload.createdBy = user.email;
                const newId = (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36));
                payload.id = newId;
            } else {
                payload.id = id;
                const existing = opcrEntries.find(x => x.id === id) || {};
                payload.createdAt = existing.createdAt || payload.updatedAt;
                payload.createdBy = existing.createdBy || user.email;
            }

            try {
                await opcrService.save(payload);
                elements.opcrModal?.hide();
                resetOpcrForm();

                // Open the form in a new tab after successful save
                const formUrl = `opcr-form.html?id=${payload.id}`;
                window.open(formUrl, '_blank');
            } catch (err) {
                console.error(err);
                alert('Failed to save OPCR entry: ' + err.message);
            }
        }

        // Wire new buttons
        function attachFormListeners() {
            const btnBack = document.getElementById('btnBackToOpcrDash');
            if (btnBack) btnBack.addEventListener('click', switchToOpcrList);

            const btnAddRow = document.getElementById('btnAddOpcrRow');
            if (btnAddRow) btnAddRow.addEventListener('click', () => addOpcrRow());

            const btnSave = document.getElementById('btnSaveOpcrForm');
            if (btnSave) btnSave.addEventListener('click', saveOpcrForm);

            // Expose globally
            window.switchToOpcrList = switchToOpcrList;
            window.saveOpcrForm = saveOpcrForm;
            window.addOpcrRow = addOpcrRow;
        }

        // Call this inside init/attachListeners
        // Re-wiring listeners

        function attachListeners() {
            if (elements.addOpcrBtn) {
                elements.addOpcrBtn.addEventListener('click', () => {
                    resetOpcrForm();
                    elements.opcrModal?.show();
                });
            }

            if (elements.opcrForm) {
                elements.opcrForm.addEventListener('submit', handleOpcrSubmit);
            }

            if (elements.opcrSearchInput) {
                elements.opcrSearchInput.addEventListener('input', () => renderOpcr());
            }

            if (elements.opcrYearFilter) {
                elements.opcrYearFilter.addEventListener('change', () => renderOpcr());
            }

            const updateStatusBtn = document.getElementById('updateOpcrStatusBtn');
            if (updateStatusBtn) {
                updateStatusBtn.addEventListener('click', () => window.saveOpcrStatusFromModal());
            }

            attachFormListeners();
        }

        // Expose functions globally for onclick handlers
        window.editOpcr = editOpcr;
        window.deleteOpcr = deleteOpcr;

        attachListeners();

        return {
            setOpcrEntries,
            getOpcrEntries,
            renderOpcr
        };
    }

    window.OpcrFeature = { init };

})(window);
