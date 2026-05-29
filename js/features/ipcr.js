// IPCR feature module: handles CRUD, rendering, and search for IPCR entries.
// Depends on IPCRService, Firebase auth, and DOM elements passed in.
(function (window) {
    if (window.IpcrFeature) return;

    function init(opts) {
        const {
            elements,
            ipcrService,
            opcrService,
            auth,
            utils,
            isAdminRef = () => false,
        } = opts || {};

        if (!elements || !ipcrService || !auth || !utils) {
            throw new Error('IpcrFeature.init missing required dependencies');
        }

        const { serverTimestamp } = window.AppFirebase || {};

        let ipcrEntries = [];

        function setIpcrEntries(list) {
            ipcrEntries = Array.isArray(list) ? list.slice() : [];
        }

        function getIpcrEntries() {
            return ipcrEntries.slice();
        }

        // ---- Rendering ----

        function ipcrRowHtml(entry) {
            const user = auth.currentUser;
            const isAdmin = isAdminRef();
            const isOwner = user && (entry.createdBy === user.email);
            const canEdit = isOwner || isAdmin;

            let actionsHtml = '';
            // Edit button - only for owner or admin
            if (canEdit) {
                actionsHtml += `<button class="btn btn-sm btn-outline-success me-1" title="Open IPCR Form" onclick="event.stopPropagation(); editIpcr('${entry.id}')"><i class="fa fa-file-lines"></i></button>`;
            }

            // Delete button - for owner or admin
            if (isAdmin || isOwner) {
                actionsHtml += `<button class="btn btn-sm btn-outline-danger" title="Delete" onclick="event.stopPropagation(); deleteIpcr('${entry.id}')"><i class="fa fa-trash"></i></button>`;
            }

            // Format timestamp
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

            // Period display
            const period = (entry.periodStart && entry.periodEnd)
                ? `${entry.periodStart.substring(0, 3)} \u2013 ${entry.periodEnd.substring(0, 3)}`
                : '\u2014';

            // Final average
            const avg = entry.finalAverage || '\u2014';

            return `<tr data-id="${entry.id}" style="cursor:pointer;" onclick="showIpcrDetailsModal('${entry.id}')" title="Click to view details">
        <td>
          <div class="fw-semibold">${entry.rateeName || entry.department || '\u2014'}</div>
          ${entry.rateePosition ? `<div class="small text-muted">${entry.rateePosition}</div>` : ''}
          ${entry.department ? `<div class="small text-muted"><i class="fa fa-building me-1"></i>${entry.department}</div>` : ''}
        </td>
        <td>
          <div>${updatedStr}</div>
          <span class="badge ${statusClass}">${currentStatus}</span>
        </td>
        <td class="text-center">${entry.year || '\u2014'}</td>
        <td class="text-center small">${period}</td>
        <td class="text-center fw-bold ${avg !== '\u2014' ? 'text-success' : 'text-muted'}">${avg}</td>
        <td onclick="event.stopPropagation();">${actionsHtml}</td>
      </tr>`;
        }

        // Show IPCR details modal
        window.showIpcrDetailsModal = function (entryId) {
            const entry = ipcrEntries.find(e => e.id === entryId);
            if (!entry) return;

            const modal = document.getElementById('ipcrDetailsModal');
            if (!modal) {
                console.error('IPCR Details Modal not found');
                return;
            }

            // Populate modal fields
            const nameEl = document.getElementById('ipcrModalRateeName');
            if (nameEl) nameEl.textContent = entry.rateeName || entry.department || 'Unknown';

            const yearEl = document.getElementById('ipcrModalYear');
            if (yearEl) yearEl.textContent = entry.year || '';

            const deptEl = document.getElementById('ipcrModalDepartment');
            if (deptEl) deptEl.textContent = entry.department || '\u2014';

            const periodEl = document.getElementById('ipcrModalPeriod');
            if (periodEl) periodEl.textContent = (entry.periodStart && entry.periodEnd)
                ? `${entry.periodStart} \u2013 ${entry.periodEnd}`
                : '\u2014';

            const avgEl = document.getElementById('ipcrModalAvg');
            if (avgEl) avgEl.textContent = entry.finalAverage || '\u2014';

            // Status dropdown
            const currentStatus = entry.status || 'Draft';
            const statusSelect = document.getElementById('ipcrModalStatus');
            if (statusSelect) {
                statusSelect.value = currentStatus;
                statusSelect.dataset.entryId = entryId;
            }

            // Edit history / versions
            const logsContainer = document.getElementById('ipcrModalLogs');
            if (logsContainer) {
                const versions = entry.versions || [];
                let historyHtml = '';

                if (versions.length > 0) {
                    const sorted = [...versions].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
                    sorted.forEach((ver, idx) => {
                        const verTime = ver.timestamp ? new Date(ver.timestamp).toLocaleString() : '';
                        const isLatest = idx === 0;
                        const badge = isLatest
                            ? `<span class="badge bg-info ms-2">Current</span>`
                            : `<span class="badge bg-secondary ms-2">v${sorted.length - idx}</span>`;

                        historyHtml += `
                                <div class="d-flex justify-content-between align-items-start py-2 border-bottom">
                                    <div class="flex-grow-1">
                                        <strong>${ver.action || 'Form saved'}</strong>
                                        <div class="small text-muted">by ${ver.user || 'Unknown'}</div>
                                    </div>
                                    <div class="text-end d-flex align-items-center">
                                        <span class="small text-muted">${verTime}</span>
                                        ${badge}
                                    </div>
                                </div>
                            `;
                    });
                }

                logsContainer.innerHTML = historyHtml || '<div class="text-muted text-center py-3"><i class="fa fa-info-circle me-1"></i>No edit history available</div>';
            }

            const bsModal = new bootstrap.Modal(modal);
            bsModal.show();
        };

        // Save status from modal
        window.saveIpcrStatusFromModal = async function () {
            const statusSelect = document.getElementById('ipcrModalStatus');
            if (!statusSelect) return;

            const entryId = statusSelect.dataset.entryId;
            const newStatus = statusSelect.value;

            try {
                const user = auth.currentUser;
                if (!user) { alert('You must be logged in.'); return; }

                await ipcrService.update(entryId, {
                    status: newStatus,
                    updatedAt: serverTimestamp(),
                    updatedBy: user.email,
                });

                const modal = bootstrap.Modal.getInstance(document.getElementById('ipcrDetailsModal'));
                if (modal) modal.hide();
            } catch (err) {
                console.error('Error updating IPCR status:', err);
                alert('Failed to update status: ' + err.message);
            }
        };

        function renderIpcr() {
            const tbody = elements.ipcrTbody;
            const empty = elements.ipcrEmpty;
            if (!tbody) return;

            let list = ipcrEntries.slice();

            // Filter by Search
            const search = (elements.ipcrSearchInput?.value || '').toLowerCase();
            if (search) {
                list = list.filter(e =>
                    (e.rateeName || '').toLowerCase().includes(search) ||
                    (e.rateePosition || '').toLowerCase().includes(search) ||
                    (e.department || '').toLowerCase().includes(search) ||
                    (e.status || '').toLowerCase().includes(search)
                );
            }

            // Filter by Year
            const yearFilter = elements.ipcrYearFilter?.value;
            if (yearFilter && yearFilter !== '') {
                list = list.filter(e => String(e.year) === String(yearFilter));
            }

            // Sort: newest year first, then by name
            list.sort((a, b) => {
                if (b.year !== a.year) return b.year - a.year;
                return (a.rateeName || '').localeCompare(b.rateeName || '');
            });

            if (list.length === 0) {
                tbody.innerHTML = '';
                if (empty) empty.style.display = 'block';
            } else {
                if (empty) empty.style.display = 'none';
                tbody.innerHTML = list.map(ipcrRowHtml).join('');
            }

            populateIpcrYearOptions();
        }

        function populateIpcrYearOptions() {
            const yearSelect = elements.ipcrYearFilter;
            if (!yearSelect) return;

            const years = new Set();
            const currentYear = new Date().getFullYear();
            years.add(currentYear);
            years.add(currentYear + 1);
            ipcrEntries.forEach(e => { if (e.year) years.add(Number(e.year)); });

            const sortedYears = Array.from(years).sort((a, b) => b - a);
            const currentVal = yearSelect.value;

            if (yearSelect.options.length <= 1) {
                yearSelect.innerHTML = '<option value="">All</option>' +
                    sortedYears.map(y => `<option value="${y}">${y}</option>`).join('');
                yearSelect.value = currentVal;
            }
        }

        // ---- CRUD ----

        let currentIpcrId = null;

        function editIpcr(id) {
            const formUrl = `ipcr-form.html?id=${id}`;
            window.open(formUrl, '_blank');
        }

        async function deleteIpcr(id) {
            const entry = ipcrEntries.find(e => e.id === id);
            const user = auth.currentUser;
            const isAdmin = isAdminRef();
            const isOwner = entry && user && entry.createdBy === user.email;

            if (!isAdmin && !isOwner) {
                alert('Only the creator of this entry or administrators can delete it.');
                return;
            }
            if (!confirm('Are you sure you want to delete this IPCR entry?')) return;
            try {
                await ipcrService.remove(id);
            } catch (err) {
                console.error(err);
                alert('Failed to delete IPCR entry: ' + err.message);
            }
        }

        async function handleIpcrSubmit(e) {
            e.preventDefault();
            const fd = new FormData(elements.ipcrForm);
            const id = fd.get('ipcrId');

            const rateeName = document.getElementById('ipcrFormRateeName')?.value || '';
            const rateePosition = document.getElementById('ipcrFormRateePosition')?.value || '';
            const department = document.getElementById('ipcrFormDepartment')?.value || '';
            const year = Number(document.getElementById('ipcrFormYear')?.value) || new Date().getFullYear();
            const periodStart = document.getElementById('ipcrFormPeriodStart')?.value || 'JANUARY';
            const periodEnd = document.getElementById('ipcrFormPeriodEnd')?.value || 'DECEMBER';
            const status = document.getElementById('ipcrFormStatus')?.value || 'Draft';
            const raterName = document.getElementById('ipcrFormRaterName')?.value || '';
            const endorserName = document.getElementById('ipcrFormEndorserName')?.value || '';
            const approverName = document.getElementById('ipcrFormApproverName')?.value || '';
            const remarks = document.getElementById('ipcrFormRemarks')?.value || '';

            const user = auth.currentUser;
            if (!user) { alert('You must be logged in.'); return; }

            const payload = {
                rateeName,
                rateePosition,
                department,
                year,
                periodStart,
                periodEnd,
                status,
                raterName,
                endorserName,
                approverName,
                remarks,
                updatedAt: serverTimestamp ? serverTimestamp() : new Date(),
                updatedBy: user.email,
            };

            if (!id) {
                payload.createdAt = serverTimestamp ? serverTimestamp() : new Date();
                payload.createdBy = user.email;
                const newId = (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36));
                payload.id = newId;
            } else {
                payload.id = id;
                const existing = ipcrEntries.find(x => x.id === id) || {};
                payload.createdAt = existing.createdAt || payload.updatedAt;
                payload.createdBy = existing.createdBy || user.email;
            }

            try {
                await ipcrService.save(payload);
                elements.ipcrModal?.hide();
                resetIpcrForm();

                // Open the form in a new tab after successful save
                const formUrl = `ipcr-form.html?id=${payload.id}`;
                window.open(formUrl, '_blank');
            } catch (err) {
                console.error(err);
                alert('Failed to save IPCR entry: ' + err.message);
            }
        }

        function resetIpcrForm() {
            if (elements.ipcrForm) elements.ipcrForm.reset();
            if (elements.ipcrId) elements.ipcrId.value = '';
            const yearInput = document.getElementById('ipcrFormYear');
            if (yearInput) yearInput.value = new Date().getFullYear();
        }

        function attachListeners() {
            if (elements.addIpcrBtn) {
                elements.addIpcrBtn.addEventListener('click', () => {
                    resetIpcrForm();
                    elements.ipcrModal?.show();
                });
            }

            if (elements.ipcrForm) {
                elements.ipcrForm.addEventListener('submit', handleIpcrSubmit);
            }

            if (elements.ipcrSearchInput) {
                elements.ipcrSearchInput.addEventListener('input', () => renderIpcr());
            }

            if (elements.ipcrYearFilter) {
                elements.ipcrYearFilter.addEventListener('change', () => renderIpcr());
            }

            const updateStatusBtn = document.getElementById('updateIpcrStatusBtn');
            if (updateStatusBtn) {
                updateStatusBtn.addEventListener('click', () => window.saveIpcrStatusFromModal());
            }

            // Generate IPCR button
            if (elements.generateIpcrBtn) {
                elements.generateIpcrBtn.addEventListener('click', () => {
                    openGenerateIpcrWizard();
                });
            }
        }

        async function copyIpcrEntry(entryId) {
            const entry = ipcrEntries.find(e => e.id === entryId);
            if (!entry) return;

            if (!confirm('Are you sure you want to duplicate this IPCR entry? This will create a new draft copy.')) return;

            try {
                const user = auth.currentUser;
                if (!user) { alert('You must be logged in.'); return; }

                const newId = (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36));
                const payload = {
                    ...entry,
                    id: newId,
                    status: 'Draft',
                    periodStart: 'JULY',
                    periodEnd: 'DECEMBER',
                    createdAt: serverTimestamp ? serverTimestamp() : new Date(),
                    createdBy: user.email,
                    updatedAt: serverTimestamp ? serverTimestamp() : new Date(),
                    updatedBy: user.email,
                    versions: [] // Reset history for the new copy
                };

                // Add a log entry
                payload.versions.push({
                    action: 'Created from copy of ' + (entry.year || ''),
                    user: user.email,
                    timestamp: new Date().toISOString()
                });

                await ipcrService.save(payload);

                // Hide details modal if open
                try {
                    const modalEl = document.getElementById('ipcrDetailsModal');
                    const bsModal = bootstrap.Modal.getInstance(modalEl);
                    if (bsModal) bsModal.hide();
                } catch (e) { }

                // Open the new form in a new tab
                const formUrl = `ipcr-form.html?id=${payload.id}`;
                window.open(formUrl, '_blank');
            } catch (err) {
                console.error('Error copying IPCR:', err);
                alert('Failed to copy IPCR: ' + err.message);
            }
        }

        // ================================================================
        // Generate IPCR from OPCR Wizard
        // ================================================================
        let genState = {
            step: 1,
            opcrList: [],
            deptMap: {},
            selectedOpcr: null,
            selectedEmployee: null,
            filteredItems: [],
        };

        function openGenerateIpcrWizard() {
            // Reset state
            genState = {
                step: 1,
                opcrList: [],
                deptMap: {},
                selectedOpcr: null,
                selectedEmployee: null,
                filteredItems: [],
            };
            showGenStep(1);
            elements.generateIpcrModal?.show();
            loadOpcrDepartments();
        }

        function showGenStep(step) {
            genState.step = step;
            // Hide all steps
            document.querySelectorAll('.gen-step').forEach(el => el.style.display = 'none');
            const stepEl = document.getElementById('genStep' + step);
            if (stepEl) stepEl.style.display = 'block';

            // Update step badges
            document.querySelectorAll('.gen-step-badge').forEach(badge => {
                const badgeStep = parseInt(badge.dataset.step);
                if (badgeStep <= step) {
                    badge.classList.remove('bg-secondary');
                    badge.classList.add('bg-info', 'text-white');
                } else {
                    badge.classList.remove('bg-info', 'text-white');
                    badge.classList.add('bg-secondary');
                }
            });

            // Show/hide buttons
            const backBtn = document.getElementById('genBackBtn');
            const createBtn = document.getElementById('genCreateBtn');
            if (backBtn) backBtn.style.display = step > 1 ? 'inline-block' : 'none';
            if (createBtn) createBtn.style.display = step === 3 ? 'inline-block' : 'none';
        }

        async function loadOpcrDepartments() {
            const loading = document.getElementById('genDeptLoading');
            const listEl = document.getElementById('genDeptList');
            const emptyEl = document.getElementById('genDeptEmpty');

            if (loading) loading.style.display = 'block';
            if (listEl) { listEl.style.display = 'none'; listEl.innerHTML = ''; }
            if (emptyEl) emptyEl.style.display = 'none';

            try {
                // Fetch all OPCRs
                const snapshot = await opcrService.collection().get();
                const opcrs = [];
                snapshot.forEach(function (doc) {
                    opcrs.push({ id: doc.id, ...doc.data() });
                });

                genState.opcrList = opcrs;

                if (loading) loading.style.display = 'none';

                if (opcrs.length === 0) {
                    if (emptyEl) emptyEl.style.display = 'block';
                    return;
                }

                // Group by department
                const deptMap = {};
                opcrs.forEach(function (opcr) {
                    const dept = opcr.department || 'Unknown';
                    if (!deptMap[dept]) {
                        deptMap[dept] = [];
                    }
                    deptMap[dept].push(opcr);
                });
                genState.deptMap = deptMap;

                // Sort departments alphabetically
                const departments = Object.keys(deptMap).sort();

                let html = '';
                departments.forEach(function (dept) {
                    const deptOpcrs = deptMap[dept].sort(function (a, b) { return (b.year || 0) - (a.year || 0); });
                    const latest = deptOpcrs[0];
                    const employeeCount = (latest.employees && Array.isArray(latest.employees)) ? latest.employees.length : 0;
                    const opcrCount = deptOpcrs.length;
                    const status = latest.status || 'Draft';
                    const statusClass = status === 'Approved' ? 'bg-success' : status === 'Submitted' ? 'bg-primary' : 'bg-secondary';

                    const deptEncoded = dept.replace(/'/g, "\\'");

                    html += '<button type="button" class="list-group-item list-group-item-action py-3" onclick="window._genSelectDept(\'' + deptEncoded + '\')">' +
                        '<div class="d-flex justify-content-between align-items-start">' +
                        '<div>' +
                        '<div class="fw-semibold"><i class="fa-solid fa-building me-2 text-info"></i>' + dept + '</div>' +
                        '<small class="text-muted">Latest: ' + (latest.year || '\u2014') + ' | ' + (latest.periodStart || '') + ' \u2013 ' + (latest.periodEnd || '') + '</small>' +
                        '<div class="mt-1"><small class="text-muted"><i class="fa-solid fa-users me-1"></i>' + employeeCount + ' employee(s) registered</small></div>' +
                        '</div>' +
                        '<div class="text-end">' +
                        '<span class="badge ' + statusClass + '">' + status + '</span>' +
                        (opcrCount > 1 ? '<br><small class="text-muted"><i class="fa-solid fa-layer-group me-1"></i>' + opcrCount + ' OPCR(s)</small>' : '') +
                        '</div>' +
                        '</div>' +
                        '</button>';
                });

                if (listEl) {
                    listEl.innerHTML = html;
                    listEl.style.display = 'block';
                }
            } catch (err) {
                console.error('Error loading OPCR departments:', err);
                if (loading) loading.style.display = 'none';
                if (emptyEl) {
                    emptyEl.innerHTML = '<i class="fa-solid fa-exclamation-triangle fa-2x mb-2 d-block text-danger"></i> Error loading OPCR data: ' + err.message;
                    emptyEl.style.display = 'block';
                }
            }
        }

        // Step 1: Select department — if multiple OPCRs, show year picker; if one, go to Step 2
        window._genSelectDept = function (deptName) {
            var deptOpcrs = genState.deptMap[deptName];
            if (!deptOpcrs || deptOpcrs.length === 0) return;

            // If only one OPCR for this department, select it directly
            if (deptOpcrs.length === 1) {
                window._genSelectOpcr(deptOpcrs[0].id);
                return;
            }

            // Multiple OPCRs — show sub-list with year selection
            var listEl = document.getElementById('genDeptList');
            if (!listEl) return;

            var html = '';
            // Back-to-departments link
            html += '<button type="button" class="list-group-item list-group-item-action py-2 text-info fw-semibold" onclick="window._genBackToDepts()">' +
                '<i class="fa-solid fa-arrow-left me-2"></i>Back to Departments' +
                '</button>';

            // Department header
            html += '<div class="list-group-item bg-light py-2">' +
                '<div class="fw-semibold"><i class="fa-solid fa-building me-2 text-info"></i>' + deptName + '</div>' +
                '<small class="text-muted">Select which OPCR year to use as source:</small>' +
                '</div>';

            // Sort by year descending
            var sorted = deptOpcrs.slice().sort(function (a, b) { return (b.year || 0) - (a.year || 0); });

            sorted.forEach(function (opcr) {
                var st = opcr.status || 'Draft';
                var stClass = st === 'Approved' ? 'bg-success' : st === 'Submitted' ? 'bg-primary' : 'bg-secondary';
                var empCount = (opcr.employees && Array.isArray(opcr.employees)) ? opcr.employees.length : 0;
                var itemCount = (opcr.items && Array.isArray(opcr.items)) ? opcr.items.filter(function (i) { return i.type === 'row'; }).length : 0;

                html += '<button type="button" class="list-group-item list-group-item-action py-3" onclick="window._genSelectOpcr(\'' + opcr.id + '\')">' +
                    '<div class="d-flex justify-content-between align-items-center">' +
                    '<div>' +
                    '<div class="fw-semibold"><i class="fa-solid fa-calendar-alt me-2 text-primary"></i>OPCR ' + (opcr.year || '\u2014') + '</div>' +
                    '<small class="text-muted">' + (opcr.periodStart || '') + ' \u2013 ' + (opcr.periodEnd || '') + '</small>' +
                    '<div class="mt-1">' +
                    '<small class="text-muted"><i class="fa-solid fa-users me-1"></i>' + empCount + ' employee(s)</small>' +
                    '<small class="text-muted ms-3"><i class="fa-solid fa-list-check me-1"></i>' + itemCount + ' indicator(s)</small>' +
                    '</div>' +
                    '</div>' +
                    '<div class="text-end">' +
                    '<span class="badge ' + stClass + ' mb-1">' + st + '</span>' +
                    '<br><i class="fa-solid fa-chevron-right text-muted"></i>' +
                    '</div>' +
                    '</div>' +
                    '</button>';
            });

            listEl.innerHTML = html;
        };

        // Go back to department list from OPCR year selection
        window._genBackToDepts = function () {
            loadOpcrDepartments();
        };

        // Select a specific OPCR and proceed to Step 2 (Employee selection)
        window._genSelectOpcr = function (opcrId) {
            var opcr = genState.opcrList.find(function (o) { return o.id === opcrId; });
            if (!opcr) return;

            genState.selectedOpcr = opcr;

            // Update labels — include the OPCR year
            var deptLabel = document.getElementById('genSelectedDept');
            if (deptLabel) deptLabel.textContent = (opcr.department || 'Unknown') + ' (OPCR ' + (opcr.year || '') + ')';

            // Set year from OPCR
            var yearInput = document.getElementById('genIpcrYear');
            if (yearInput) yearInput.value = opcr.year || new Date().getFullYear();

            // Set period from OPCR
            var periodStartSel = document.getElementById('genIpcrPeriodStart');
            var periodEndSel = document.getElementById('genIpcrPeriodEnd');
            if (periodStartSel && opcr.periodStart) periodStartSel.value = opcr.periodStart;
            if (periodEndSel && opcr.periodEnd) periodEndSel.value = opcr.periodEnd;

            showGenStep(2);
            loadOpcrEmployees(opcr);
        };

        function loadOpcrEmployees(opcr) {
            const listEl = document.getElementById('genEmpList');
            const emptyEl = document.getElementById('genEmpEmpty');

            if (listEl) listEl.innerHTML = '';
            if (emptyEl) emptyEl.style.display = 'none';

            const employees = (opcr.employees && Array.isArray(opcr.employees))
                ? opcr.employees.filter(function (emp) { return emp.name && emp.name.trim() !== ''; })
                : [];

            if (employees.length === 0) {
                if (emptyEl) emptyEl.style.display = 'block';
                return;
            }

            let html = '';
            employees.forEach(function (emp, idx) {
                var initials = (emp.abbreviation || emp.name.charAt(0) || '?').toUpperCase().substring(0, 2);
                html += '<button type="button" class="list-group-item list-group-item-action py-2" onclick="window._genSelectEmployee(' + idx + ')">' +
                    '<div class="d-flex align-items-center">' +
                    '<div class="rounded-circle bg-info text-white d-flex align-items-center justify-content-center me-3" style="width:36px;height:36px;font-size:14px;font-weight:600;">' +
                    initials +
                    '</div>' +
                    '<div>' +
                    '<div class="fw-semibold">' + emp.name + '</div>' +
                    (emp.abbreviation ? '<small class="text-muted">Abbreviation: ' + emp.abbreviation + '</small>' : '') +
                    '</div>' +
                    '<i class="fa-solid fa-chevron-right ms-auto text-muted"></i>' +
                    '</div>' +
                    '</button>';
            });

            if (listEl) {
                listEl.innerHTML = html;
            }
        }

        // Step 2 -> Step 3: Select employee and build preview
        window._genSelectEmployee = function (empIndex) {
            const opcr = genState.selectedOpcr;
            if (!opcr || !opcr.employees) return;

            const employee = opcr.employees[empIndex];
            if (!employee) return;

            genState.selectedEmployee = employee;

            // Update labels
            const empLabel = document.getElementById('genSelectedEmp');
            const deptLabel2 = document.getElementById('genSelectedDept2');
            if (empLabel) empLabel.textContent = employee.name;
            if (deptLabel2) deptLabel2.textContent = opcr.department || 'Unknown';

            showGenStep(3);
            buildPreviewItems(opcr, employee);
        };

        function buildPreviewItems(opcr, employee) {
            const items = opcr.items || [];
            const previewTbody = document.getElementById('genPreviewTbody');
            const emptyEl = document.getElementById('genPreviewEmpty');
            const selectAllCb = document.getElementById('genSelectAll');

            if (!previewTbody) return;
            previewTbody.innerHTML = '';
            if (emptyEl) emptyEl.style.display = 'none';

            if (items.length === 0) {
                if (emptyEl) emptyEl.style.display = 'block';
                genState.filteredItems = [];
                return;
            }

            const empName = (employee.name || '').trim().toLowerCase();
            const empAbbr = (employee.abbreviation || '').trim().toLowerCase();

            // Build enriched items with match info
            const enrichedItems = items.map(function (item, idx) {
                var isMatch = false;
                var matchReason = '';

                if (item.type === 'row') {
                    var responsible = (item.responsible || '').toLowerCase();

                    // Check if employee name or abbreviation appears in responsible field
                    if (empName && responsible.includes(empName)) {
                        isMatch = true;
                        matchReason = 'Name match';
                    } else if (empAbbr && empAbbr.length > 0 && responsible.includes(empAbbr)) {
                        isMatch = true;
                        matchReason = 'Abbreviation match';
                    }

                    // Check for department-wide keywords like "FOMD personnel", "all personnel", "all staff", etc.
                    // Also match "All" or "all" as a standalone word (applies to all employees)
                    var deptKeywords = ['personnel', 'all staff', 'all employee', 'all personnel', 'division personnel', 'all'];
                    if (!isMatch) {
                        for (var k = 0; k < deptKeywords.length; k++) {
                            var kw = deptKeywords[k];
                            if (kw === 'all') {
                                // Strict whole-word check to avoid matching names like "Allen"
                                var words = responsible.split(/[\s,/]+/);
                                if (words.indexOf('all') !== -1) {
                                    isMatch = true;
                                    matchReason = 'All Employees';
                                    break;
                                }
                            } else if (responsible.includes(kw)) {
                                isMatch = true;
                                matchReason = 'Department-wide (' + responsible.trim() + ')';
                                break;
                            }
                        }
                    }
                } else {
                    // Section/subsection headers - include them to preserve structure
                    isMatch = true;
                    matchReason = 'Structure';
                }

                return Object.assign({}, item, { _index: idx, _isMatch: isMatch, _matchReason: matchReason });
            });

            genState.filteredItems = enrichedItems;

            // Render preview
            var html = '';
            enrichedItems.forEach(function (item, idx) {
                var typeLabel = getItemTypeLabel(item.type);
                var typeBadgeClass = getItemTypeBadgeClass(item.type);
                var content = getItemContent(item);
                var checked = item._isMatch ? 'checked' : '';
                var rowClass = item._isMatch ? '' : 'table-light text-muted';

                html += '<tr class="' + rowClass + '">' +
                    '<td class="text-center">' +
                    '<input type="checkbox" class="gen-item-cb" data-idx="' + idx + '" ' + checked + '>' +
                    '</td>' +
                    '<td>' +
                    '<span class="badge ' + typeBadgeClass + ' small">' + typeLabel + '</span>' +
                    (item._matchReason && item.type === 'row' ? '<br><small class="text-success"><i class="fa-solid fa-check-circle"></i> ' + item._matchReason + '</small>' : '') +
                    '</td>' +
                    '<td class="small">' + content + '</td>' +
                    '</tr>';
            });

            previewTbody.innerHTML = html;

            // Select all checkbox
            if (selectAllCb) {
                selectAllCb.checked = true;
                selectAllCb.onchange = function () {
                    var cbs = previewTbody.querySelectorAll('.gen-item-cb');
                    cbs.forEach(function (cb) { cb.checked = selectAllCb.checked; });
                };
            }
        }

        function getItemTypeLabel(type) {
            switch (type) {
                case 'section': return 'Section';
                case 'subsection1': return 'MFO';
                case 'subsection2': return 'PI';
                case 'row': return 'Data';
                default: return type || 'Unknown';
            }
        }

        function getItemTypeBadgeClass(type) {
            switch (type) {
                case 'section': return 'bg-dark';
                case 'subsection1': return 'bg-info text-white';
                case 'subsection2': return 'bg-warning text-dark';
                case 'row': return 'bg-light text-dark border';
                default: return 'bg-secondary';
            }
        }

        function getItemContent(item) {
            if (item.type === 'section' || item.type === 'subsection1' || item.type === 'subsection2') {
                return '<strong>' + (item.title || '(untitled)') + '</strong>';
            }
            if (item.type === 'row') {
                var indicator = (item.indicator || '').substring(0, 100);
                var targets = (item.targets || '').substring(0, 80);
                var responsible = item.responsible || '';
                return indicator + (targets ? ' | <em>Targets: ' + targets + '</em>' : '') + (responsible ? ' | <small class="text-muted">Resp: ' + responsible + '</small>' : '');
            }
            return item.title || item.indicator || '\u2014';
        }

        // ---- Month-to-Quarter mapping for evidence link extraction ----
        const MONTH_ORDER = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];

        function monthIndex(name) {
            return MONTH_ORDER.indexOf((name || '').toUpperCase());
        }

        /**
         * Given an IPCR period (startMonth, endMonth), return which OPCR quarters
         * overlap with that period.  Returns an array like ['Q1','Q2'].
         */
        function getQuartersForPeriod(startMonth, endMonth) {
            const s = monthIndex(startMonth);
            const e = monthIndex(endMonth);
            if (s < 0 || e < 0) return ['Q1','Q2','Q3','Q4']; // fallback: all

            const quarters = [];
            // Q1 = Jan(0)–Mar(2), Q2 = Apr(3)–Jun(5), Q3 = Jul(6)–Sep(8), Q4 = Oct(9)–Dec(11)
            const qRanges = [
                { name: 'Q1', start: 0, end: 2 },
                { name: 'Q2', start: 3, end: 5 },
                { name: 'Q3', start: 6, end: 8 },
                { name: 'Q4', start: 9, end: 11 }
            ];
            qRanges.forEach(function (q) {
                // Check if quarter range overlaps with [s, e]
                if (q.start <= e && q.end >= s) {
                    quarters.push(q.name);
                }
            });
            return quarters.length > 0 ? quarters : ['Q1','Q2','Q3','Q4'];
        }

        /**
         * Extract and merge evidence links & accomplishment text from an OPCR row
         * for the given quarters.
         */
        function extractEvidenceForQuarters(opcrRow, quarters) {
            const allLinks = [];
            const accTexts = [];
            const seenUrls = new Set();

            quarters.forEach(function (q) {
                const qLower = q.toLowerCase(); // q1, q2, ...
                // Evidence links
                const linksKey = 'links' + q; // linksQ1, linksQ2, ...
                const links = opcrRow[linksKey];
                if (Array.isArray(links)) {
                    links.forEach(function (link) {
                        if (link && link.url && !seenUrls.has(link.url)) {
                            seenUrls.add(link.url);
                            allLinks.push({ url: link.url, label: link.label || 'Evidence', quarter: q });
                        }
                    });
                }
                // Accomplishment text
                const accKey = 'acc' + q; // accQ1, accQ2, ...
                const accText = (opcrRow[accKey] || '').trim();
                if (accText) {
                    accTexts.push(q + ': ' + accText);
                }
            });

            return { evidenceLinks: allLinks, accomplishmentText: accTexts.join('\n') };
        }

        // Generate IPCR from selected items
        async function generateIpcrFromOpcr() {
            const user = auth.currentUser;
            if (!user) { alert('You must be logged in.'); return; }

            const opcr = genState.selectedOpcr;
            const employee = genState.selectedEmployee;
            if (!opcr || !employee) {
                alert('Invalid selection. Please try again.');
                return;
            }

            // Collect checked items
            const previewTbody = document.getElementById('genPreviewTbody');
            const checkedIndices = new Set();
            if (previewTbody) {
                previewTbody.querySelectorAll('.gen-item-cb:checked').forEach(function (cb) {
                    checkedIndices.add(parseInt(cb.dataset.idx));
                });
            }

            // Determine which OPCR quarters fall within the IPCR period
            const ipcrPeriodStart = document.getElementById('genIpcrPeriodStart')?.value || opcr.periodStart || 'JANUARY';
            const ipcrPeriodEnd = document.getElementById('genIpcrPeriodEnd')?.value || opcr.periodEnd || 'DECEMBER';
            const matchingQuarters = getQuartersForPeriod(ipcrPeriodStart, ipcrPeriodEnd);
            console.log('IPCR period:', ipcrPeriodStart, '-', ipcrPeriodEnd, '=> Matching quarters:', matchingQuarters);

            // Build the IPCR items from selected OPCR items
            const opcrItems = genState.filteredItems || [];
            const ipcrItems = [];

            opcrItems.forEach(function (item, idx) {
                if (!checkedIndices.has(idx)) return;

                if (item.type === 'section') {
                    ipcrItems.push({ type: 'section', title: item.title || '' });
                } else if (item.type === 'subsection1') {
                    ipcrItems.push({ type: 'subsection1', title: item.title || '' });
                } else if (item.type === 'subsection2') {
                    ipcrItems.push({ type: 'subsection2', title: item.title || '' });
                } else if (item.type === 'row') {
                    // Extract evidence links & accomplishments from matching quarters
                    var evidence = extractEvidenceForQuarters(item, matchingQuarters);

                    // Map OPCR row to IPCR data row — include evidence links
                    ipcrItems.push({
                        type: 'data',
                        mfo: item.indicator || '',
                        targets: item.targets || '',
                        accomplishments: evidence.accomplishmentText || '',
                        ratingQ: '',
                        ratingE: '',
                        ratingT: '',
                        ratingAve: '',
                        remarks: '',
                        evidenceLinks: evidence.evidenceLinks || []
                    });
                }
            });

            // If no items selected, add default structure
            if (ipcrItems.length === 0) {
                ipcrItems.push(
                    { type: 'section', title: 'A. Major Final Outputs (MFOs)/ Operations' },
                    { type: 'subsection1', title: 'MFO1: Strategic Planning and Management' },
                    { type: 'subsection2', title: 'Performance Indicator 1' },
                    { type: 'data', mfo: '', targets: '', accomplishments: '', ratingQ: '', ratingE: '', ratingT: '', ratingAve: '', remarks: '' }
                );
            }

            // Remove orphan section/subsection headers
            const cleanedItems = removeOrphanHeaders(ipcrItems);

            const year = Number(document.getElementById('genIpcrYear')?.value) || opcr.year || new Date().getFullYear();
            const periodStart = document.getElementById('genIpcrPeriodStart')?.value || opcr.periodStart || 'JANUARY';
            const periodEnd = document.getElementById('genIpcrPeriodEnd')?.value || opcr.periodEnd || 'DECEMBER';

            const newId = crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36);

            // Collect signatory names from wizard
            const sigDeptManager = (document.getElementById('genSigDeptManager')?.value || '').trim();
            const sigDeputyAdmin = (document.getElementById('genSigDeputyAdmin')?.value || '').trim();
            const sigAdmin = (document.getElementById('genSigAdmin')?.value || '').trim();

            const payload = {
                id: newId,
                rateeName: employee.name || '',
                rateePosition: '',
                department: opcr.department || '',
                year: year,
                periodStart: periodStart,
                periodEnd: periodEnd,
                status: 'Draft',
                raterName: sigDeptManager,
                endorserName: sigDeputyAdmin,
                approverName: sigAdmin,
                // Signatory fields used by ipcr-form.html
                sigRaterName: sigDeptManager,
                sigRaterPos: sigDeptManager ? 'Manager, ' + (opcr.department || '') : '',
                sigReviewerName: sigDeptManager,
                sigRecommenderName: sigDeputyAdmin,
                sigApproverName: sigAdmin,
                remarks: 'Generated from OPCR: ' + (opcr.department || '') + ' (' + (opcr.year || '') + ')',
                items: cleanedItems,
                sourceOpcrId: opcr.id,
                createdAt: serverTimestamp ? serverTimestamp() : new Date(),
                createdBy: user.email,
                updatedAt: serverTimestamp ? serverTimestamp() : new Date(),
                updatedBy: user.email,
            };

            try {
                const createBtn = document.getElementById('genCreateBtn');
                if (createBtn) {
                    createBtn.disabled = true;
                    createBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin me-1"></i>Generating...';
                }

                await ipcrService.save(payload);

                elements.generateIpcrModal?.hide();

                const formUrl = 'ipcr-form.html?id=' + newId;
                window.open(formUrl, '_blank');

                console.log('IPCR generated successfully from OPCR:', opcr.department, '->', employee.name);
            } catch (err) {
                console.error('Error generating IPCR:', err);
                alert('Failed to generate IPCR: ' + err.message);
            } finally {
                const createBtn = document.getElementById('genCreateBtn');
                if (createBtn) {
                    createBtn.disabled = false;
                    createBtn.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles me-1"></i>Generate IPCR';
                }
            }
        }

        // Remove orphan section/subsection headers that have no data rows after them
        function removeOrphanHeaders(items) {
            const result = [];
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (item.type === 'section' || item.type === 'subsection1' || item.type === 'subsection2') {
                    var hasDataAfter = false;
                    for (var j = i + 1; j < items.length; j++) {
                        if (items[j].type === 'data') {
                            hasDataAfter = true;
                            break;
                        }
                        if (item.type === 'section' && items[j].type === 'section') break;
                        if (item.type === 'subsection1' && (items[j].type === 'section' || items[j].type === 'subsection1')) break;
                        if (item.type === 'subsection2' && (items[j].type === 'section' || items[j].type === 'subsection1' || items[j].type === 'subsection2')) break;
                    }
                    if (hasDataAfter) {
                        result.push(item);
                    }
                } else {
                    result.push(item);
                }
            }
            return result;
        }

        // Wire up wizard navigation buttons
        (function attachGenListeners() {
            var backBtn = document.getElementById('genBackBtn');
            var createBtn = document.getElementById('genCreateBtn');

            if (backBtn) {
                backBtn.addEventListener('click', function () {
                    if (genState.step === 2) {
                        showGenStep(1);
                    } else if (genState.step === 3) {
                        showGenStep(2);
                    }
                });
            }

            if (createBtn) {
                createBtn.addEventListener('click', function () {
                    generateIpcrFromOpcr();
                });
            }
        })();

        // Expose globally for onclick handlers
        window.editIpcr = editIpcr;
        window.deleteIpcr = deleteIpcr;
        window.copyIpcrEntry = copyIpcrEntry;

        attachListeners();

        return {
            setIpcrEntries,
            getIpcrEntries,
            renderIpcr,
        };
    }

    window.IpcrFeature = { init };

})(window);
