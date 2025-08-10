/*
 * Construction Project Monitoring Dashboard Script
 * Handles CRUD operations, filtering, exporting, and rendering charts
 * Data persistence uses localStorage for demo purposes.
 */

(() => {
  const LS_KEY = "construction_projects"; // deprecated, kept for backward compatibility
  const db = firebase.firestore();
  // Enable verbose Firestore console logging for debugging
  firebase.firestore.setLogLevel('debug');
  const PROJECTS_COL = 'projects';
  const DEEPWELLS_COL = 'deepwells';
  const REFORESTATION_COL = 'reforestations';
  let unsubscribeProjects = null;
  let unsubscribeDeepwells = null;
  let unsubscribeReforestations = null;

  const ADMIN_EMAIL = "johnlowel.fradejas@mwss.gov.ph"; // change to your address
  const ADMIN_EMAIL_LOWER = ADMIN_EMAIL.toLowerCase();
  let isAdmin = false;

  const elements = {
    projectsContainer: document.getElementById("projectsContainer"),
    agencyFilter: document.getElementById("agencyFilter"),
    statusFilter: document.getElementById("statusFilter"),
    searchInput: document.getElementById("searchInput"),
    projectForm: document.getElementById("projectForm"),
    projectModal: (function(){ const el = document.getElementById('projectModal'); return el ? new bootstrap.Modal(el) : { show(){}, hide(){} }; })(),
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
    dwProviderFilter: document.getElementById('dwProviderFilter'),
    dwStatusFilter: document.getElementById('dwStatusFilter'),
    dwSearchInput: document.getElementById('dwSearchInput'),
    deepwellForm: document.getElementById('deepwellForm'),
    deepwellModal: (function(){ const el = document.getElementById('deepwellModal'); return el ? new bootstrap.Modal(el) : { show(){}, hide(){} }; })(),
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
    reforestationModal: (function(){ const el = document.getElementById('reforestationModal'); return el ? new bootstrap.Modal(el) : { show(){}, hide(){} }; })(),
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
    reforestationDetailsModal: (function(){ const el = document.getElementById('reforestationDetailsModal'); return el ? new bootstrap.Modal(el) : { show(){}, hide(){} }; })(),
  };

  // convenient references
  const {loginForm, signupForm, showSignupLink, showLoginLink, loginLinks, signupLinks} = elements;
  const appWrapper = document.getElementById('appWrapper');
  const billingContainer = document.getElementById('billingContainer');
  const addBillingBtn    = document.getElementById('addBillingBtn');
  const viewOnlyLink = document.getElementById('viewOnlyLink');
  const logoutBtn  = document.getElementById('logoutBtn');
  const loginBtn   = document.getElementById('loginBtn');
  const loginScreen = document.getElementById('loginScreen');

  let projects = [];
  let deepwells = [];
  let reforestations = [];
  let isViewOnly = false;
  let isLevel2 = false; // accessLevel 2 users
  let elevatedAccess = false; // isAdmin || isLevel2
  let pendingUsers = [];
  let dwMonthlyChart = null; // Chart.js instance for deepwells monthly totals
  let approvedUsers = [];
  function updatePendingBadge(){
    const badge = document.getElementById('pendingBadge');
    if(!badge) return;
    const cnt = pendingUsers.length;
    if(cnt>0){badge.textContent=cnt;badge.style.display='inline-block';}
    else {badge.style.display='none';}
  }
  

  // Load config for signups
  async function loadConfig(){
    try{
      const doc = await db.collection('config').doc('app').get();
      if(doc.exists){
        const data = doc.data();
        // signup enable/disable config deprecated
      }
    }catch(err){console.warn('Failed to load config',err);} 
    updateAdminUI();
      subscribeDeepwells();
  }

  function updateAdminUI(){
    if(isAdmin){
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

  async function loadPendingUsers(){
  try{
    const snap = await db.collection('users').where('approved','==',false).get();
    pendingUsers = snap.docs.map(d=>({id:d.id,...d.data()}));
    renderPendingUsersTable();
    updatePendingBadge();
  }catch(err){console.error('Failed to load pending users',err);}
}

// Load and render approved users list
async function loadApprovedUsers(){
  try{
    const snap = await db.collection('users').where('approved','==',true).get();
    approvedUsers = snap.docs.map(d=>({id:d.id,...d.data()}));
    renderApprovedUsersTable();
      if(window.__refreshChatRecipients) window.__refreshChatRecipients();
  }catch(err){console.error('Failed to load approved users',err);}  
}

function renderApprovedUsersTable(){
  if(!elements || !elements.approvedUsersTable){ return; }
  const tbody = elements.approvedUsersTable.querySelector('tbody');
  tbody.innerHTML='';
  approvedUsers.forEach(u=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${u.email}</td><td>${u.updatedAt?.toDate?u.updatedAt.toDate().toLocaleString():''}</td><td></td>`;
    const levelCell = tr.lastElementChild;
    const select = document.createElement('select');
    select.className = 'form-select form-select-sm';
    [1,2].forEach(l=>{
      const opt=document.createElement('option');
      opt.value=l;
      opt.textContent = 'Level '+l;
      if((u.accessLevel||1)===l) opt.selected=true;
      select.appendChild(opt);
    });
    select.onchange = () => updateUserAccessLevel(u, parseInt(select.value,10));
    levelCell.appendChild(select);
    tbody.appendChild(tr);
  });
}

async function updateUserAccessLevel(u, level){
  try{
    await db.collection('users').doc(u.id).update({accessLevel:level});
    u.accessLevel = level;
  }catch(err){alert('Failed to update access level: '+err.message);}  
}

function renderPendingUsersTable(){
  const tbody = elements.pendingUsersTable.querySelector('tbody');
  tbody.innerHTML='';
  pendingUsers.forEach(u=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${u.email}</td><td>${u.createdAt?.toDate?u.createdAt.toDate().toLocaleString():''}</td><td></td>`;
    const actions=tr.lastElementChild;
    const approve=document.createElement('button');
    approve.className='btn btn-sm btn-success me-1';
    approve.textContent='Approve';
    approve.onclick=()=>approveUser(u);
    const reject=document.createElement('button');
    reject.className='btn btn-sm btn-danger';
    reject.textContent='Reject';
    reject.onclick=()=>rejectUser(u);
    actions.append(approve,reject);
    tbody.appendChild(tr);
  });
}

async function approveUser(u){
  try{
    await db.collection('users').doc(u.id).update({approved:true, accessLevel:1});
    firebase.auth().sendPasswordResetEmail(u.email).catch(()=>{});
    loadPendingUsers();
  }catch(e){alert(e.message);} 
}
async function rejectUser(u){
  try{
    await db.collection('users').doc(u.id).delete();
    loadPendingUsers();
  }catch(e){alert(e.message);} 
}



  // ---- Billing Details handlers ----
  function createBillingRow(data={}){
    const div=document.createElement('div');
    div.className='row g-2 align-items-end mb-2 billing-row';
    div.innerHTML=`<div class="col-4"><input type="date" class="form-control billing-date" value="${data.date||''}" placeholder="Date"/></div>
                  <div class="col-4"><input type="number" class="form-control billing-amount" value="${data.amount||''}" placeholder="Amount (PHP)" step="0.01" min="0"/></div>
                  <div class="col-3"><input type="text" class="form-control billing-desc" value="${data.desc||''}" placeholder="Description"/></div>
                  <div class="col-1 text-end"><button type="button" class="btn btn-outline-danger btn-sm remove-billing"><i class="fa fa-minus"></i></button></div>`;
    div.querySelector('.remove-billing').onclick=()=>{
      div.remove();
    };
    billingContainer.appendChild(div);
  }
  if(addBillingBtn){
    addBillingBtn.addEventListener('click',()=>createBillingRow());
  }

  function gatherBilling(){
    const rows = Array.from(billingContainer.querySelectorAll('.billing-row'));
    return rows.map(r=>({
      date:  r.querySelector('.billing-date').value,
      amount: parseFloat(r.querySelector('.billing-amount').value)||0,
      desc: r.querySelector('.billing-desc').value.trim()
    })).filter(b=>b.date && b.amount);
  }
  function populateBilling(arr){
    billingContainer.innerHTML='';
    (arr||[]).forEach(d=>createBillingRow(d));
  }

  // ---- Project CRUD & Rendering ----

  async function saveProject(project) {
    await db.collection(PROJECTS_COL).doc(project.id).set(project);
  }

  function onSaveProject(e) {
    e.preventDefault();
    const formData = new FormData(elements.projectForm);
  const progressBilling = gatherBilling();
    const id = formData.get("projectId") || (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).substr(2, 5));
    if(!formData.get("projectName").trim()){alert('Project Name is required');return;}
    if(!formData.get("contractor").trim()){alert('Contractor is required');return;}
    const accDateRaw=formData.get("accompDate");
    const accDate = accDateRaw || new Date().toISOString().slice(0,10);
    const currentPercent=parseFloat(formData.get("percentAccomplishment"))||0;
    const prevPercent=parseFloat(formData.get("percentPrevious"))||0;
    const plannedPercent=parseFloat(formData.get("percentPlanned"))||0;
    const variance=+(currentPercent-plannedPercent).toFixed(2);
    const newAcc={date:accDate,percent:currentPercent,prevPercent,plannedPercent,variance,activities:formData.get("activities"),issue:formData.get("issues"),action:formData.get("actionTaken"),remarks:formData.get("remarks")};
    const existing = projects.find(p=>p.id===id)||{};
    const project={
      id,
      name:formData.get("projectName"),
      implementingAgency:formData.get("implementingAgency"),
    location: formData.get("projectLocation"),
      contractor:formData.get("contractor"),
      contractAmount:parseFloat(formData.get("contractAmount"))||0,
      revisedContractAmount:formData.get("revisedContractAmount")?parseFloat(formData.get("revisedContractAmount")):null,
      contractDocsLink: elevatedAccess ? (formData.get("contractDocsLink")?.trim() || '') : (existing.contractDocsLink || ''),
      ntpDate:formData.get("ntpDate"),
      originalDuration:parseInt(formData.get("originalDuration"),10)||0,
      timeExtension:parseInt(formData.get("timeExtension"),10)||0,
      originalCompletion:formData.get("originalCompletion"),
      revisedCompletion:formData.get("revisedCompletion"),
      activities:formData.get("activities"),
      issues:formData.get("issues"),
      remarks:formData.get("remarks"),
      otherDetails: formData.get("otherDetails"),
      progressBilling,
      history:[...(existing.history||[])],
      photos: existing.photos || (existing.sCurveDataUrl ? [existing.sCurveDataUrl] : []),
      accomplishments:[...(existing.accomplishments||[])]
    };
    const userEmail=firebase.auth().currentUser?.email||'unknown';
    if(!project.history) project.history=[];
    project.history.push({email:userEmail,timestamp:new Date().toISOString(),action:formData.get("projectId")?'edit':'create'});
    if(newAcc.date){
      const idx = project.accomplishments.findIndex(a=>a.date===newAcc.date);
      if(idx>-1){
        project.accomplishments[idx] = {...newAcc}; // overwrite existing entry for the same date
      }else{
        project.accomplishments.push({...newAcc});
      }
    }
    function postSaveUI(){
      // Update local cache immediately so UI reflects latest changes before Firestore snapshot returns
      const idxLocal = projects.findIndex(p => p.id === project.id);
      if (idxLocal > -1) {
        projects[idxLocal] = { ...project };
      } else {
        projects.push({ ...project });
      }
      render(); // refresh list UI

      // reset filters & form
      elements.searchInput.value = '';
      elements.agencyFilter.value = '';
      elements.statusFilter.value = '';
      elements.projectModal.hide();
      clearForm();
      billingContainer.innerHTML = '';
    }
    const photoInputs=["projectPhoto1","projectPhoto2","projectPhoto3"].map(id=>document.getElementById(id));
    const files=photoInputs.map(inp=>inp?.files[0]).filter(f=>!!f);
    if(files.length){
      Promise.all(files.map(f=>compressImage(f,1024,0.75)))
        .then(dataUrls=>{
          // ensure JPEG data URLs at a consistent max dimension
          project.photos = dataUrls.slice(0,3);
          return saveProject(project);
        })
        .then(postSaveUI)
        .catch(err=>alert(err.message));
    }else{
      saveProject(project).then(postSaveUI).catch(err=>alert(err.message));
    }
  }

  window.deleteProject = async function(id){ if(!elevatedAccess){alert('Only admin/level2 can delete projects');return;} try{await db.collection(PROJECTS_COL).doc(id).delete();projects=projects.filter(p=>p.id!==id);render();}catch(err){alert(err.message);} };

  function clearForm(){
  billingContainer.innerHTML='';elements.projectForm.reset();document.getElementById("projectId").value="";["projectPhoto1","projectPhoto2","projectPhoto3"].forEach(id=>{const el=document.getElementById(id);if(el) el.value="";});const linkInput=document.getElementById("contractDocsLink");const linkGroup=document.getElementById("contractDocsGroup");linkGroup.style.display = elevatedAccess ? '' : 'none';linkInput.value="";linkInput.disabled=!elevatedAccess;document.getElementById("actionTaken").value="";document.getElementById("percentAccomplishment").value=0;document.getElementById("percentPrevious").value=0;document.getElementById("percentPlanned").value=0;document.getElementById("accompDate").value="";}

  function populateForm(project){
  populateBilling(project.progressBilling);document.getElementById("projectId").value=project.id;document.getElementById("projectName").value=project.name;document.getElementById("implementingAgency").value=project.implementingAgency;
document.getElementById("projectLocation").value=project.location || '';document.getElementById("contractor").value=project.contractor;document.getElementById("contractAmount").value=project.contractAmount;document.getElementById("revisedContractAmount").value=project.revisedContractAmount??'';
const linkGroup=document.getElementById("contractDocsGroup");const linkInput=document.getElementById("contractDocsLink");linkGroup.style.display = elevatedAccess ? '' : 'none';
linkInput.value = project.contractDocsLink || '';
linkInput.disabled = !elevatedAccess;document.getElementById("ntpDate").value=project.ntpDate;document.getElementById("originalDuration").value=project.originalDuration;document.getElementById("timeExtension").value=project.timeExtension;document.getElementById("originalCompletion").value=project.originalCompletion;document.getElementById("revisedCompletion").value=project.revisedCompletion;document.getElementById("activities").value=project.activities;document.getElementById("issues").value=project.issues;document.getElementById("actionTaken").value="";document.getElementById("remarks").value=project.remarks;document.getElementById("otherDetails").value=project.otherDetails;const last=(project.accomplishments||[]).slice(-1)[0]||{percent:0,prevPercent:0,date:""};document.getElementById("percentPrevious").value=last.prevPercent??last.percent??0;document.getElementById("percentPlanned").value=last.plannedPercent??0;document.getElementById("percentAccomplishment").value=last.percent??0;document.getElementById("accompDate").value=last.date;document.getElementById("actionTaken").value=last.action||"";}

  function getProjectStatus(p){
      // Get the most recent accomplishment by date, not by array position
      const accomplishments = p.accomplishments || [];
      const latest = accomplishments.length > 0 
        ? accomplishments.slice().sort((a, b) => new Date(b.date) - new Date(a.date))[0]
        : null;
      if(latest){
        // Completed if cumulative accomplishment has reached 100 %
        if((latest.percent ?? 0) >= 100) return 'Completed';
        // If actual to-date accomplishment lags behind planned, mark as Delayed
        if((latest.percent ?? 0) < (latest.plannedPercent ?? 0)) return 'Delayed';
      }
      // Fallback to schedule-based check
      const today = new Date().toISOString().split('T')[0];
      if(p.revisedCompletion && today > p.revisedCompletion) return 'Delayed';
      if(!p.revisedCompletion && today > p.originalCompletion) return 'Delayed';
      return 'On-going';
    }

  function projectRowHtml(p){
    // Get the most recent accomplishment by date, not by array position
    const accomplishments = p.accomplishments || [];
    const latest = accomplishments.length > 0 
      ? accomplishments.slice().sort((a, b) => new Date(b.date) - new Date(a.date))[0]
      : { percent: 0 };
    const curr = latest.percent ?? 0;
    let actionsHtml = (!isViewOnly && firebase.auth().currentUser) ? `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="editProject('${p.id}')"><i class="fa fa-pencil"></i></button>` : '';
    if(isAdmin){
      actionsHtml = `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="editProject('${p.id}')"><i class="fa fa-pencil"></i></button><button class="btn btn-sm btn-danger" title="Delete" onclick="deleteProject('${p.id}')"><i class="fa fa-trash"></i></button>`;
    }
    return `<tr data-id="${p.id}"><td>${p.name}</td><td>${p.implementingAgency || ''}</td><td>${p.contractor}</td><td><span class="badge bg-${getProjectStatus(p) === 'Delayed' ? 'danger' : 'primary'}">${getProjectStatus(p)}</span></td><td>${p.revisedCompletion || p.originalCompletion}</td><td>${curr}%</td><td class="d-flex gap-1">${actionsHtml}</td></tr>`;
  }

  function render(){const addBtn=document.getElementById('addProjectBtn');if(addBtn) addBtn.style.display=(!isViewOnly && firebase.auth().currentUser)?'inline-block':'none';const agencies = [...new Set(projects.map(p => p.implementingAgency))];
      const prevAgency = elements.agencyFilter.value;
      // rebuild options only if count changed (or first render)
      if(elements.agencyFilter.options.length - 1 !== agencies.length){
        elements.agencyFilter.innerHTML = `<option value="">All Agencies</option>` + agencies.map(a => `<option value="${a}">${a}</option>`).join("");
      }
      // restore previous selection if still in list
      if(prevAgency && agencies.includes(prevAgency)){
        elements.agencyFilter.value = prevAgency;
      }
      const text = elements.searchInput.value.toLowerCase();
      const agency = elements.agencyFilter.value;const status = elements.statusFilter.value;
      const filtered = projects.filter(p => {
        const projStatus = getProjectStatus(p);
        const matchText = (p.name || "").toLowerCase().includes(text) || (p.contractor || "").toLowerCase().includes(text);
        const matchAgency = agency ? p.implementingAgency === agency : true;
        let matchStatus = true;
        if(status){
          if(status === 'On-going'){
            // treat both On-going and Delayed as active work
            matchStatus = projStatus === 'On-going' || projStatus === 'Delayed';
          }else{
            matchStatus = projStatus === status;
          }
        }
        return matchText && matchAgency && matchStatus;
      });const tbody=document.getElementById('projectsTbody');tbody.innerHTML=filtered.map(projectRowHtml).join('');attachRowEvents();}

  function attachRowEvents(){const tbody=document.getElementById('projectsTbody');Array.from(tbody.querySelectorAll('tr')).forEach(row=>{row.addEventListener('click',e=>{if(isViewOnly) return; if(e.target.closest('button')) return;const pid=row.dataset.id;viewProject(pid);});});}

  window.editProject=(id)=>{const p=projects.find(x=>x.id===id);if(!p) return;populateForm(p);elements.projectModal.show();};

  window.viewProject=(id)=>{const p=projects.find(proj=>proj.id===id);if(!p) return;const photos=(p.photos&&p.photos.length)?p.photos:(p.sCurveDataUrl?[p.sCurveDataUrl]:[]);
  const billingHtml = (p.progressBilling&&p.progressBilling.length)?`<h6 class="mt-3">Billing Details</h6><div class="table-responsive"><table class="table table-sm"><thead><tr><th>Date</th><th>Amount (PHP)</th><th>Description</th></tr></thead><tbody>${p.progressBilling.map(b=>`<tr><td>${b.date}</td><td>₱${Number(b.amount).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td><td>${b.desc||''}</td></tr>`).join('')}</tbody></table></div>`:'';const photosHtml=photos.length?`<div class="photo-grid mb-3">`+photos.slice(0,3).map((url,i)=>`<div class="photo-tile" aria-label="Project photo ${i+1}"><img src="${url}" data-full="${url}" data-index="${i}" alt="Project photo ${i+1}" class="photo-img" loading="lazy"></div>`).join('')+`</div>`:'';const accompSorted = (p.accomplishments || []).slice().sort((a,b)=> new Date(b.date) - new Date(a.date));
    // Remove near-duplicate rows (same date & key fields)
    const seenKeys = new Set();
    const accompDedup = accompSorted.filter(a=>{
      const key = `${a.date}|${a.percent}|${a.prevPercent}|${a.plannedPercent}|${a.activities||''}|${a.issue||''}|${a.action||''}|${a.remarks||''}`;
      if(seenKeys.has(key)) return false;
      seenKeys.add(key);
      return true;
    });
    const editHistorySorted = (p.history || []).slice().sort((a,b)=> new Date(b.timestamp) - new Date(a.timestamp));
    const body = document.getElementById('detailsBody');body.innerHTML=`<h5 class=\"fw-bold mb-2\">${p.name}</h5><p><strong>Implementing Agency:</strong> ${p.implementingAgency}</p><p><strong>Contractor:</strong> ${p.contractor}</p>
<p><strong>Location:</strong> ${p.location || ''}</p><p><strong>Contract Amount:</strong> ₱${Number(p.contractAmount||0).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</p>${p.revisedContractAmount?`<p><strong>Revised Contract Amount:</strong> ₱${Number(p.revisedContractAmount).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</p>`:''}<p><strong>Status:</strong> ${getProjectStatus(p)}</p><p><strong>NTP:</strong> ${p.ntpDate}</p><p><strong>Duration:</strong> ${p.originalDuration} days ${p.timeExtension?"+"+p.timeExtension:""}</p><p><strong>Target Completion:</strong> ${p.revisedCompletion||p.originalCompletion}</p>${p.otherDetails?`<p><strong>Details:</strong> ${p.otherDetails}</p>`:''}${p.contractDocsLink?`<p><strong>Contract Docs:</strong> ${elevatedAccess?`<a href="${p.contractDocsLink}" target="_blank">Open</a>`:`<span class="text-muted">No authority to access</span>`}</p>`:''}<h6 class="mt-3">Accomplishment History</h6><div class="table-responsive"><table class="table table-sm"><thead><tr><th>Date</th><th class="percent-col">Planned %</th><th class="percent-col">Previous %</th><th class="percent-col">To Date %</th><th class="percent-col">Variance %</th><th class="w-25">Activities</th><th class="w-20">Issue</th><th class="w-20">Action Taken</th><th class="w-25">Remarks</th></tr></thead><tbody>${accompDedup.map(a=>`<tr><td class="date-col">${a.date}</td><td class="percent-col">${(a.plannedPercent??0).toFixed(2)}%</td><td class="percent-col">${(a.prevPercent??0).toFixed(2)}%</td><td class="percent-col">${(a.percent??0).toFixed(2)}%</td><td class="percent-col">${a.variance>=0?'+':''}${(a.variance??0).toFixed(2)}%</td><td class="w-25">${bulletizeActivities(a.activities)}</td><td class="w-20">${a.issue||p.issues||''}</td><td class="w-20">${a.action||''}</td><td class="w-25">${a.remarks||''}</td></tr>`).join('')}</tbody></table></div>${billingHtml}${editHistorySorted.length?`<h6 class="mt-3">Edit History</h6><div class="table-responsive"><table class="table table-sm"><thead><tr><th>User</th><th>Timestamp</th><th>Action</th></tr></thead><tbody>${editHistorySorted.map(h=>`<tr><td>${h.email}</td><td>${new Date(h.timestamp).toLocaleString()}</td><td>${h.action}</td></tr>`).join('')}</tbody></table></div>`:''}`;body.innerHTML+=photosHtml;const list=photos.slice();body.querySelectorAll('.photo-img').forEach(img=>{img.style.cursor='zoom-in';img.addEventListener('click',()=>{const idx=parseInt(img.getAttribute('data-index'))||0;window.openPhotoViewer(idx,list);});});const footer=document.getElementById('detailsFooter');footer.innerHTML=`<button class="btn btn-outline-success" id="docxBtn"><i class="fa fa-download me-1"></i>Word</button>`;document.getElementById('docxBtn').onclick=()=>exportProjectDocx(p);bootstrap.Modal.getOrCreateInstance(document.getElementById('detailsModal')).show();};

  // --- Photo Viewer helpers ---
  const viewerState = { list: [], index: 0, open: false };

  function showViewerImage(){
    const imgEl = document.getElementById('photoViewerImg');
    if(imgEl) imgEl.src = viewerState.list[viewerState.index] || '';
  }
  function updateViewerNav(){
    const prevBtn = document.getElementById('photoPrev');
    const nextBtn = document.getElementById('photoNext');
    const multi = viewerState.list.length > 1;
    if(prevBtn) prevBtn.style.display = multi ? 'flex' : 'none';
    if(nextBtn) nextBtn.style.display = multi ? 'flex' : 'none';
  }
  function prevPhoto(){
    if(viewerState.list.length<=1) return;
    viewerState.index = (viewerState.index - 1 + viewerState.list.length) % viewerState.list.length;
    showViewerImage();
  }
  function nextPhoto(){
    if(viewerState.list.length<=1) return;
    viewerState.index = (viewerState.index + 1) % viewerState.list.length;
    showViewerImage();
  }

  window.openPhotoViewer = function(indexOrSrc, list){
    try{
      const modalEl = document.getElementById('photoViewer');
      if(!modalEl) return;
      if(Array.isArray(list)){
        viewerState.list = list.slice();
        const n = parseInt(indexOrSrc);
        viewerState.index = isNaN(n)?0:Math.max(0, Math.min(n, viewerState.list.length-1));
      }else{
        const src = typeof indexOrSrc === 'string' ? indexOrSrc : '';
        viewerState.list = [src];
        viewerState.index = 0;
      }
      showViewerImage();
      updateViewerNav();
      bootstrap.Modal.getOrCreateInstance(modalEl).show();
    }catch(e){ console.error('openPhotoViewer error', e); }
  };

  // Setup controls and keyboard handling
  document.addEventListener('DOMContentLoaded', ()=>{
    const modalEl = document.getElementById('photoViewer');
    const imgEl = document.getElementById('photoViewerImg');
    const prevBtn = document.getElementById('photoPrev');
    const nextBtn = document.getElementById('photoNext');
    if(!modalEl || !imgEl) return;
    // Click image to close
    imgEl.addEventListener('click', ()=> bootstrap.Modal.getOrCreateInstance(modalEl).hide());
    // Prev/Next buttons
    if(prevBtn) prevBtn.addEventListener('click', e=>{ e.stopPropagation(); prevPhoto(); });
    if(nextBtn) nextBtn.addEventListener('click', e=>{ e.stopPropagation(); nextPhoto(); });

    // Keyboard
    const onKeyDown = (e)=>{
      if(!viewerState.open) return;
      if(e.key==='ArrowLeft'){ e.preventDefault(); prevPhoto(); }
      else if(e.key==='ArrowRight'){ e.preventDefault(); nextPhoto(); }
    };

    modalEl.addEventListener('shown.bs.modal', ()=>{
      viewerState.open = true;
      document.addEventListener('keydown', onKeyDown);
    });
    modalEl.addEventListener('hidden.bs.modal', ()=>{
      viewerState.open = false;
      document.removeEventListener('keydown', onKeyDown);
      if(imgEl) imgEl.src = '';
      viewerState.list = [];
      viewerState.index = 0;
    });
  });

  // simple docx export skipped for brevity ...

// ----- UI Cleanup Helpers -----
function cleanDeepwellDuplicates(){
  const dwSec = document.getElementById('deepwellsSection');
  if(!dwSec) return;
  // Remove duplicate search/filter controls (keep first occurrence)
  const searchInputs = dwSec.querySelectorAll('#dwSearchInput');
  searchInputs.forEach((el,idx)=>{ if(idx>0) el.closest('.col-md-3, .col-md-2, .col-md-3.text-md-end')?.remove();});
  const providerFilters = dwSec.querySelectorAll('#dwProviderFilter');
  providerFilters.forEach((el,idx)=>{ if(idx>0) el.closest('.col-md-2')?.remove();});
  const statusFilters = dwSec.querySelectorAll('#dwStatusFilter');
  statusFilters.forEach((el,idx)=>{ if(idx>0) el.closest('.col-md-2')?.remove();});
  const addBtns = dwSec.querySelectorAll('#addDeepwellBtn');
  addBtns.forEach((btn,idx)=>{ if(idx>0) btn.closest('.col-md-3')?.remove();});
  // Remove duplicate tables (keep first)
  const tables = dwSec.querySelectorAll('#deepwellsTable');
  tables.forEach((tbl,idx)=>{ if(idx>0) tbl.closest('.table-responsive')?.remove();});
}

document.addEventListener('DOMContentLoaded',cleanDeepwellDuplicates);


/* ----------------- Facebook Messenger Feature ----------------- */
(function(){
  const messengerToggle = document.getElementById('messengerToggle');
  const messengerSidebar = document.getElementById('messengerSidebar');
  const messengerClose = document.getElementById('messengerClose');
  const messengerBadge = document.getElementById('messengerBadge');
  const msgContainer = document.getElementById('chatMessages');
  const msgInput = document.getElementById('chatInput');
  const msgSend = document.getElementById('chatSend');
  const contactsList = document.getElementById('contactsList');
  const messengerClear = document.getElementById('messengerClear');
  
  let currentChat = 'all';
  let currentChatTitle = 'General Chat';
  let currentChatAvatar = 'fa-users';
  
  // updateBadge is defined later in this scope (single consolidated implementation)
  // Update General Chat badge in the contacts header
  function updateGeneralBadge(){
    const badge = document.querySelector('.contact-item[data-chat="all"] .contact-badge');
    if(!badge) return;
    const count = unreadCounts['all']||0;
    badge.textContent = count;
    badge.style.display = count>0 ? 'inline-flex' : 'none';
  }
  // Toast container for notifications
  const toastContainer = document.createElement('div');
  toastContainer.className = 'toast-container position-fixed bottom-0 end-0 p-3';
  toastContainer.style.zIndex = '1080';
  document.body.appendChild(toastContainer);
  
  function notifyPm(fromEmail){
    // Toast pop-up disabled at user's request; we still update the unread badge
    if(!fromEmail) return;
    updateBadge();
  }
  
  // Messenger UI Functions
  function showMessenger(){
    messengerSidebar.style.display = 'flex';
    document.body.style.overflow = 'hidden';
    // Clear unread for the currently open chat and refresh view
    try { switchChat(currentChat, currentChatTitle, currentChatAvatar); } catch(_) {}
  }
  
  function hideMessenger(){
    messengerSidebar.style.display = 'none';
    document.body.style.overflow = '';
  }
  
  function switchChat(chatId, title, avatar){
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
    if(chatTitle) chatTitle.textContent = title;
    if(chatAvatar) chatAvatar.className = `fa ${avatar}`;
    
    // Clear unread count for this chat (both UID and possible email key)
    if(unreadCounts[chatId]) unreadCounts[chatId] = 0;
    if(chatId === ADMIN_EMAIL_LOWER && ADMIN_UID && unreadCounts[ADMIN_UID]) unreadCounts[ADMIN_UID]=0;
    if(chatId === 'all') unreadCounts['all']=0;
    updateBadge();
    updateContactsList();
    
    renderMessages();
  }
  // Old chat elements removed - using Facebook Messenger elements instead
  function ensureClearBtn(){
    const user = firebase.auth().currentUser;
    const email = user && user.email ? user.email.trim().toLowerCase() : '';
    const isAdminNow = email === ADMIN_EMAIL_LOWER;
    if(!isAdminNow) return;
    // Show admin clear button in messenger (null-safe)
    const clearBtn = (typeof messengerClear !== 'undefined' && messengerClear)
      ? messengerClear
      : document.getElementById('messengerClear');
    if(clearBtn){
      clearBtn.classList.remove('d-none');
    }
  }
  function attachClearHandler(){
    // Clear handler moved to Facebook Messenger implementation
    // No longer needed here
  }
  ensureClearBtn();
  // On pages without messenger UI (e.g., mobile), skip messenger init but continue app.
  if(messengerToggle){
  // Make messenger toggle draggable and persist its position
  (function initMessengerToggleDrag(){
    try{
      // Restore saved position if any
      const saved = localStorage.getItem('messengerTogglePos');
      if(saved){
        const pos = JSON.parse(saved);
        if(Number.isFinite(pos.left) && Number.isFinite(pos.top)){
          messengerToggle.style.left = pos.left + 'px';
          messengerToggle.style.top  = pos.top  + 'px';
          messengerToggle.style.right = 'auto';
          messengerToggle.style.bottom = 'auto';
        }
      }

      const drag={ active:false, moved:false, sx:0, sy:0, sl:0, st:0 };
      function onDown(ev){
        const e = ev.touches? ev.touches[0] : ev;
        drag.active=true; drag.moved=false;
        drag.sx = e.clientX; drag.sy = e.clientY;
        const rect = messengerToggle.getBoundingClientRect();
        drag.sl = rect.left; drag.st = rect.top;
        messengerToggle.style.right='auto';
        messengerToggle.style.bottom='auto';
        document.addEventListener('mousemove', onMove);
        document.addEventListener('touchmove', onMove, {passive:false});
        document.addEventListener('mouseup', onUp);
        document.addEventListener('touchend', onUp);
        if(ev.cancelable) ev.preventDefault();
      }
      function onMove(ev){
        if(!drag.active) return;
        const e = ev.touches? ev.touches[0] : ev;
        let left = drag.sl + (e.clientX - drag.sx);
        let top  = drag.st + (e.clientY - drag.sy);
        const maxLeft = window.innerWidth - messengerToggle.offsetWidth - 8;
        const maxTop  = window.innerHeight - messengerToggle.offsetHeight - 8;
        left = Math.max(8, Math.min(maxLeft, left));
        top  = Math.max(8, Math.min(maxTop, top));
        messengerToggle.style.left = left + 'px';
        messengerToggle.style.top  = top  + 'px';
        drag.moved = true;
        if(ev.cancelable) ev.preventDefault();
      }
      function onUp(){
        if(!drag.active) return;
        drag.active=false;
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('touchmove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.removeEventListener('touchend', onUp);
        const left = parseInt(messengerToggle.style.left)||0;
        const top  = parseInt(messengerToggle.style.top)||0;
        try{ localStorage.setItem('messengerTogglePos', JSON.stringify({left, top})); }catch(_){/* ignore */}
        // If it was a drag, suppress the immediate click-open
        if(drag.moved){
          messengerToggle.addEventListener('click', e=>{ e.stopImmediatePropagation(); e.preventDefault(); }, {once:true});
        }
      }
      messengerToggle.addEventListener('mousedown', onDown);
      messengerToggle.addEventListener('touchstart', onDown, {passive:false});
      window.addEventListener('resize', ()=>{
        // keep within viewport on resize
        const rect = messengerToggle.getBoundingClientRect();
        const maxLeft = window.innerWidth - messengerToggle.offsetWidth - 8;
        const maxTop  = window.innerHeight - messengerToggle.offsetHeight - 8;
        const newLeft = Math.max(8, Math.min(maxLeft, rect.left));
        const newTop  = Math.max(8, Math.min(maxTop, rect.top));
        messengerToggle.style.left = newLeft + 'px';
        messengerToggle.style.top  = newTop  + 'px';
      });
    }catch(err){ console.warn('Toggle drag init failed', err); }
  })();
  firebase.auth().onAuthStateChanged(()=>{
    ensureClearBtn();
  });
  }
  // Draggable messenger button handled by initMessengerToggleDrag()

  let msgsUnsub = null;
  const unreadCounts = {}; // {fromId: number}
  let ADMIN_UID = null;
  let allMessages = [];
  // Update main toggle badge and general chat badge
  function updateBadge(){
    // Sum all unread counts
    let total = 0;
    for(const k in unreadCounts){ total += Number(unreadCounts[k]||0); }
    if(messengerBadge){
      if(total > 0){
        messengerBadge.textContent = String(total);
        messengerBadge.style.display = 'inline-block';
      }else{
        messengerBadge.style.display = 'none';
        messengerBadge.textContent = '0';
      }
    }
    // Update General Chat badge inside contacts list
    const generalBadge = document.querySelector('.contact-item[data-chat="all"] .contact-badge');
    if(generalBadge){
      const gen = unreadCounts['all'] || 0;
      if(gen > 0){
        generalBadge.textContent = String(gen);
        generalBadge.style.display = '';
      }else{
        generalBadge.style.display = 'none';
        generalBadge.textContent = '0';
      }
    }
  }
  async function ensureAdminUid(){
    if(ADMIN_UID) return ADMIN_UID;
    try{
      const snap = await db.collection('users').where('email','==',ADMIN_EMAIL).limit(1).get();
      if(!snap.empty){
        ADMIN_UID = snap.docs[0].id;
      }
    }catch(err){console.error('Failed to fetch admin uid',err);}  
    return ADMIN_UID;
  }

  // Old chat button functions removed - using Facebook Messenger functions instead

  // populate recipients now handled by updateContactsList(); legacy refreshRecipients removed
  
  // Render Facebook-style message bubble
  function appendMsg(m, canDelete=false){
    const self = firebase.auth().currentUser?.uid === m.fromId;
    const messageGroup = document.createElement('div');
    messageGroup.className = 'message-group';
    
    // Format timestamp
    let ts = '';
    if(m.timestamp){
      const d = m.timestamp.toDate ? m.timestamp.toDate() : (m.timestamp.seconds? new Date(m.timestamp.seconds*1000) : new Date(m.timestamp));
      ts = d.toLocaleString(undefined,{hour:'2-digit',minute:'2-digit',hour12:false});
    }
    
    // Show sender name for received messages in group chats
    if(!self && currentChat === 'all'){
      const senderDiv = document.createElement('div');
      senderDiv.className = 'message-sender';
      senderDiv.textContent = m.fromEmail;
      messageGroup.appendChild(senderDiv);
    }
    
    const bubble = document.createElement('div');
    bubble.className = `message-bubble ${self ? 'sent' : 'received'}`;
    bubble.textContent = m.text;
    bubble.dataset.id = m.id || '';
    
    if(canDelete && self){
      const deleteBtn = document.createElement('button');
      deleteBtn.className = 'btn btn-sm btn-link text-danger position-absolute';
      deleteBtn.style.top = '0';
      deleteBtn.style.right = '-25px';
      deleteBtn.innerHTML = '<i class="fa fa-trash" style="font-size:10px;"></i>';
      deleteBtn.title = 'Delete message';
      deleteBtn.onclick = async () => {
        if(!confirm('Delete this message?')) return;
        deleteBtn.disabled = true;
        deleteBtn.innerHTML = '<i class="fa fa-spinner fa-spin" style="font-size:10px;"></i>';
        
        try{
          const user = firebase.auth().currentUser;
          // Check if user owns this message
          if(m.fromId !== user.uid) {
            alert('You can only delete your own messages');
            deleteBtn.disabled = false;
            deleteBtn.innerHTML = '<i class="fa fa-trash" style="font-size:10px;"></i>';
            return;
          }
          
          // Try to delete the message first
          try {
            await db.collection('messages').doc(m.id).delete();
            // Remove the message element from UI immediately
            messageGroup.remove();
          } catch(deleteErr) {
            // If direct deletion fails due to permissions, try soft delete
            console.log('Direct delete failed, trying soft delete:', deleteErr);
            await db.collection('messages').doc(m.id).update({
              deleted: true,
              deletedBy: user.uid,
              deletedAt: firebase.firestore.FieldValue.serverTimestamp()
            });
            // Remove the message element from UI immediately
            messageGroup.remove();
          }
          
        }catch(err){
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
    if(ts){
      const timeDiv = document.createElement('div');
      timeDiv.className = 'message-time';
      timeDiv.textContent = ts;
      messageGroup.appendChild(timeDiv);
    }
    
    msgContainer.appendChild(messageGroup);
    msgContainer.scrollTop = msgContainer.scrollHeight;
  }
  
  // Update contacts list with unread counts
  function updateContactsList(){
    if(!approvedUsers) return;
    
    contactsList.innerHTML = '';
    // Ensure General Chat is present at top
    (function addGeneralChat(){
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
      if(currentChat === 'all'){ genItem.classList.add('active'); }
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
      
      if(user.id === currentChat){ contactItem.classList.add('active'); }
      contactItem.addEventListener('click', () => {
        switchChat(user.id, user.email, 'fa-user');
      });
      
      contactsList.appendChild(contactItem);
    });
    
    // Add admin contact if not already present
    if(![...approvedUsers].some(u => u.email.toLowerCase() === ADMIN_EMAIL_LOWER)){
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
      
      if(ADMIN_EMAIL_LOWER === currentChat){ adminContact.classList.add('active'); }
      adminContact.addEventListener('click', () => {
        switchChat(ADMIN_EMAIL_LOWER, ADMIN_EMAIL, 'fa-crown');
      });
      
      contactsList.appendChild(adminContact);
    }
  }

  // --- Facebook Messenger filtering logic ---
  function renderMessages(){
    const uid = firebase.auth().currentUser.uid;
    const userEmail = firebase.auth().currentUser.email?.toLowerCase();
    msgContainer.innerHTML = '';
    
    allMessages.forEach(m=>{
      if(currentChat==='all'){
        if(!m.toId){ // broadcast only
          appendMsg(m, m.fromId===uid);
        }
      }else{
        // Private thread between current user and selected recipient
        const other = currentChat; // could be uid or email string
        if(m.toId===null) return; // skip broadcasts
        
        const fromSelf = m.fromId===uid;
        const fromOther = m.fromId===other || (m.fromEmail && m.fromEmail.toLowerCase()===other);
        
        // Handle both UID and email-based targeting
        const toOther = m.toId===other || (typeof m.toId==='string' && m.toId.toLowerCase && m.toId.toLowerCase()===other);
        const toMe = m.toId===uid || (m.toId && m.toId.toLowerCase && m.toId.toLowerCase()===userEmail) || (isAdmin && m.toId===ADMIN_EMAIL_LOWER && other===ADMIN_EMAIL_LOWER);
        
        const a = fromSelf && toOther; // message I sent to other
        const b = fromOther && toMe; // message other sent to me
        
        if(a || b){
          appendMsg(m, m.fromId===uid);
          // mark as read using sender UID; also clear email key if any
          if(m.fromId){ unreadCounts[m.fromId]=0; }
          if(m.fromEmail){ unreadCounts[m.fromEmail.toLowerCase()] = 0; }
          updateBadge();
          updateContactsList();
        }
      }
    });
  }
  
  // (Removed legacy single-image openPhotoViewer; using list-aware viewer defined earlier)

  function startListening(){
    if(msgsUnsub) msgsUnsub();
    const uid = firebase.auth().currentUser.uid;
    const userEmail = firebase.auth().currentUser.email?.toLowerCase();
    if(!Array.isArray(allMessages)) allMessages = [];
    let initialized = false;
    msgsUnsub = db.collection('messages').orderBy('timestamp','asc').onSnapshot(snap=>{
      // On initial snapshot, populate messages without counting unread
      if(!initialized){
        allMessages = [];
        snap.forEach(doc => {
          const m = doc.data();
          if(m.deleted) return;
          const includeForMe = (!m.toId)
            || (m.toId===uid)
            || (m.toId && m.toId.toLowerCase && m.toId.toLowerCase()===userEmail)
            || (isAdmin && m.toId===ADMIN_EMAIL_LOWER)
            || (m.fromId===uid);
          if(includeForMe){ allMessages.push({id:doc.id, ...m}); }
        });
        updateContactsList();
        renderMessages();
        initialized = true;
        return;
      }
      const changes = snap.docChanges();
      let hasAnyChange = false;
      changes.forEach(change => {
        const doc = change.doc;
        const m = doc.data();
        const id = doc.id;
        // Helper predicates
        const includeForMe = (!m.toId)
          || (m.toId===uid)
          || (m.toId && m.toId.toLowerCase && m.toId.toLowerCase()===userEmail)
          || (isAdmin && m.toId===ADMIN_EMAIL_LOWER)
          || (m.fromId===uid);
        const idx = allMessages.findIndex(mm => mm.id === id);
        if(change.type === 'removed'){
          if(idx !== -1){ allMessages.splice(idx,1); hasAnyChange = true; }
          return;
        }
        if(m.deleted){
          // treat as removed from view
          if(idx !== -1){ allMessages.splice(idx,1); hasAnyChange = true; }
          return;
        }
        if(change.type === 'modified'){
          if(idx !== -1){ allMessages[idx] = {id, ...m}; hasAnyChange = true; }
          return;
        }
        if(change.type === 'added'){
          if(includeForMe){
            allMessages.push({id, ...m});
            hasAnyChange = true;
            // Unread counts only for newly added incoming messages
            if(m.fromId !== uid){
              if(!m.toId){
                if(messengerSidebar.style.display==='none' || currentChat!=='all'){
                  unreadCounts['all'] = (unreadCounts['all']||0)+1;
                  updateBadge();
                  notifyPm(m.fromEmail);
                }
              }else{
                const messageToMe = (m.toId===uid || (m.toId && m.toId.toLowerCase && m.toId.toLowerCase()===userEmail) || (isAdmin && m.toId===ADMIN_EMAIL_LOWER));
                if(messageToMe){
                  const chatKey = m.fromId;
                  if(messengerSidebar.style.display==='none' || currentChat!==chatKey){
                    unreadCounts[chatKey] = (unreadCounts[chatKey]||0)+1;
                    updateBadge();
                    updateContactsList();
                    notifyPm(m.fromEmail);
                  }
                }
              }
              if(messengerSidebar.style.display==='none'){
                messengerToggle.classList.add('animate__animated','animate__tada');
                setTimeout(()=>messengerToggle.classList.remove('animate__animated','animate__tada'),1000);
              }
            }
          }
        }
      });
      if(hasAnyChange){
        updateContactsList();
        renderMessages();
      }
    });
  }

  // Send message function for Facebook Messenger
  async function sendMessage(){
    const txt = msgInput.value.trim();
    if(!txt) return;
    const user = firebase.auth().currentUser;
    
    // optimistic append
    appendMsg({text:txt,fromId:user.uid,fromEmail:user.email,toId:currentChat==='all'?null:currentChat,timestamp:new Date()}, true);
    msgInput.value='';
    try{ autoResize(); }catch(_){ }
    
    try{
      await db.collection('messages').add({
        text: txt,
        fromId: user.uid,
        fromEmail: user.email,
        toId: currentChat==='all'? null : currentChat,
        timestamp: firebase.firestore.FieldValue.serverTimestamp()
      });
    }catch(err){
      alert('Failed to send: '+err.message);
      console.error('sendMessage error',err);
    }
  }

  // Event handlers for Facebook Messenger (guarded for pages without messenger UI)
  if(messengerToggle){
    messengerToggle.onclick = () => {
      showMessenger();
      messengerToggle.classList.remove('animate__animated','animate__tada');
    };
  }
  if(typeof messengerClose !== 'undefined' && messengerClose){
    messengerClose.onclick = hideMessenger;
  }
  if(typeof msgSend !== 'undefined' && msgSend){
    msgSend.onclick = sendMessage;
  }
  
  if(msgInput){
    msgInput.addEventListener('keyup', e => {
      if(e.key === 'Enter' && !e.shiftKey){ e.preventDefault(); sendMessage(); }
    });
  }
  // Auto-resize textarea height
  function autoResize(){
    if(!msgInput) return;
    msgInput.style.height = 'auto';
    const max = 120; // px max height
    msgInput.style.height = Math.min(max, msgInput.scrollHeight) + 'px';
  }
  if(msgInput){
    msgInput.addEventListener('input', autoResize);
  }
  // Initialize height
  autoResize();
  
  // General chat contact click handler
  const generalContact = document.querySelector('.contact-item[data-chat="all"]');
  if(generalContact){
    generalContact.addEventListener('click', () => {
      switchChat('all', 'General Chat', 'fa-users');
    });
  }
  
  // expose for other functions
  window.__refreshChatRecipients = updateContactsList;

  // Show messenger when user logged in & approved (null-safe for pages without messenger UI)
  function showMessengerBtn(){ if(!messengerToggle) return; messengerToggle.style.display='inline-flex'; }
  function hideMessengerBtn(){ if(!messengerToggle) return; messengerToggle.style.display='none'; }

  firebase.auth().onAuthStateChanged(user=>{
    try{
      if(user && !user.isAnonymous){
        showMessengerBtn();
        if(typeof startListening === 'function'){
          startListening();
        }
        // Resolve admin UID for unread mapping and refresh contacts
        if(typeof ensureAdminUid === 'function'){
          ensureAdminUid().then(updateContactsList).catch(()=>{});
        } else {
          // Fallback: still try to refresh recipients if available
          if(typeof updateContactsList === 'function'){
            updateContactsList();
          }
        }
        // Always subscribe reforestations after login
        subscribeReforestations();
      }else{
        hideMessengerBtn();
        if(typeof msgsUnsub === 'function') msgsUnsub();
      }
    }catch(err){
      console.warn('Messenger auth listener error (ignored):', err);
      // Ensure reforestations still subscribe on login
      if(user && !user.isAnonymous){
        try{ subscribeReforestations(); }catch(_){/* ignore */}
      }
    }
  });
  
  // Admin clear messages functionality (null-safe)
  (function attachMessengerClearHandler(){
    const clearBtn = document.getElementById('messengerClear');
    if(!clearBtn) return;
    clearBtn.addEventListener('click', async ()=>{
      if(!isAdmin) return;
      if(!confirm('Delete ALL chat messages?')) return;
      clearBtn.disabled = true;
      try{
        const snap = await db.collection('messages').get();
        const batch = db.batch();
        snap.docs.forEach(doc=>batch.delete(doc.ref));
        await batch.commit();
        const msgContainer = document.getElementById('msgContainer');
        if(msgContainer) msgContainer.innerHTML='';
      }catch(err){
        alert('Failed to clear messages: '+err.message);
        console.error('Clear chat error',err);
      }finally{
        clearBtn.disabled = false;
      }
    });
  })();
})();

  /* Page Scroll & Modal Scroll Buttons */
  const scrollTopBtn=document.getElementById('scrollTopBtn');
  const scrollBottomBtn=document.getElementById('scrollBottomBtn');
  function toggleScrollButtons(){
      const scrolled = document.documentElement.scrollTop || document.body.scrollTop;
      const maxScroll = document.documentElement.scrollHeight - window.innerHeight;
      // Only show the buttons when the page is at least 200 px taller than the viewport
      if (maxScroll < 200) {
        if (scrollTopBtn) scrollTopBtn.style.display = 'none';
        if (scrollBottomBtn) scrollBottomBtn.style.display = 'none';
        return;
      }
      if (scrollTopBtn) scrollTopBtn.style.display = scrolled > 100 ? 'flex' : 'none';
      if (scrollBottomBtn) scrollBottomBtn.style.display = scrolled < maxScroll - 100 ? 'flex' : 'none';
    }
  if (scrollTopBtn && scrollBottomBtn) {
    window.addEventListener('scroll', toggleScrollButtons);
    toggleScrollButtons();
    scrollTopBtn.addEventListener('click', ()=>{
      window.scrollTo({top:0,behavior:'smooth'});
    });
    scrollBottomBtn.addEventListener('click', ()=>{
      window.scrollTo({top:document.documentElement.scrollHeight,behavior:'smooth'});
    });
  }

  // Global Modal Scroll Buttons (generic for any active modal)
  const modalUp = document.getElementById('modalScrollUp');
  const modalDown = document.getElementById('modalScrollDown');
  let activeModalBody = null;
  function getActiveModalBody(){
    const shown = document.querySelector('.modal.show');
    if(!shown) return null;
    return shown.querySelector('.modal-dialog-scrollable .modal-body') || shown.querySelector('.modal-body');
  }
  function positionModalScrollBtns(){
    if(!modalUp || !modalDown) return;
    const dlg = document.querySelector('.modal.show .modal-dialog');
    if(!dlg) return;
    const rect = dlg.getBoundingClientRect();
    const left = Math.min(rect.right - 44, window.innerWidth - 52);
    modalUp.style.left = left + 'px';
    modalDown.style.left = left + 'px';
    modalUp.style.right = 'auto';
    modalDown.style.right = 'auto';
  }
  function updateModalScrollBtns(){
    if(!modalUp || !modalDown) return;
    // Hide modal/page scroll arrows for selected modals
    const shown = document.querySelector('.modal.show');
    const suppressed = shown && ['detailsModal','deepwellModal','reforestationModal','reforestationDetailsModal'].includes(shown.id);
    if(suppressed){
      modalUp.style.display='none';
      modalDown.style.display='none';
      // Also hide global page scroll buttons to avoid overlay on details modal
      const pageTop = document.getElementById('scrollTopBtn');
      const pageBottom = document.getElementById('scrollBottomBtn');
      if(pageTop) pageTop.style.display = 'none';
      if(pageBottom) pageBottom.style.display = 'none';
      activeModalBody = null;
      return;
    }
    activeModalBody = getActiveModalBody();
    if(!activeModalBody){
      modalUp.style.display='none';
      modalDown.style.display='none';
      return;
    }
    positionModalScrollBtns();
    const scrolled = activeModalBody.scrollTop;
    const max = activeModalBody.scrollHeight - activeModalBody.clientHeight;
    modalUp.style.display = scrolled > 50 ? 'flex' : 'none';
    modalDown.style.display = scrolled < max - 50 ? 'flex' : 'none';
  }
  if(modalUp){
    modalUp.addEventListener('click',()=>{
      const body = getActiveModalBody();
      if(!body) return;
      body.scrollBy({top:-300,behavior:'smooth'});
    });
  }
  if(modalDown){
    modalDown.addEventListener('click',()=>{
      const body = getActiveModalBody();
      if(!body) return;
      body.scrollBy({top:300,behavior:'smooth'});
    });
  }
  window.addEventListener('resize', updateModalScrollBtns);
  document.addEventListener('shown.bs.modal', updateModalScrollBtns);
  document.addEventListener('hidden.bs.modal', ()=>{
    if(!modalUp || !modalDown) return;
    modalUp.style.display='none';
    modalDown.style.display='none';
    activeModalBody=null;
    // Recompute page scroll button visibility when modal closes
    if(typeof toggleScrollButtons === 'function'){
      try{ toggleScrollButtons(); }catch(_){ /* noop */ }
    }
  });
  // Capture scroll events on modal bodies to update buttons in real-time
  document.addEventListener('scroll', ()=>{
    if(document.querySelector('.modal.show')) updateModalScrollBtns();
  }, true);
  // View-only link
  if(viewOnlyLink){
    viewOnlyLink.addEventListener('click', async e=>{
      e.preventDefault();
      try{
        await firebase.auth().signInAnonymously();
      }catch(err){
        console.error('Anon sign-in failed',err);
        console.warn('Anon sign-in unavailable, proceeding without auth');
        isViewOnly = true;
        showApp();
        subscribeDeepwells();
        if(unsubscribeProjects) unsubscribeProjects();
        unsubscribeProjects = db.collection(PROJECTS_COL).onSnapshot(snap=>{
          projects = snap.docs.map(d=>({id:d.id,...d.data()}));
          render();
        });
      }
    });
  }
  // Attach listeners
  if(elements.projectForm) elements.projectForm.addEventListener('submit', onSaveProject);
  if(elements.searchInput) elements.searchInput.addEventListener('input', render);
  // Add listeners for Construction Projects filters
  elements.agencyFilter?.addEventListener('change', render);
  elements.statusFilter?.addEventListener('change', render);
// Deepwell filters
elements.dwProviderFilter?.addEventListener('change', renderDeepwells);
elements.dwStatusFilter?.addEventListener('change', renderDeepwells);
elements.dwSearchInput?.addEventListener('input', renderDeepwells);
// Reforestation filters
elements.reforestationTypeFilter?.addEventListener('change', renderReforestations);
elements.reforestationStatusFilter?.addEventListener('change', renderReforestations);
elements.reforestationSearchInput?.addEventListener('input', renderReforestations);

// ---- Utility ----
// Compress image using canvas to keep Firestore doc size small (<1MB total)
  async function compressImage(file, maxSize=1024, quality=0.7){
return new Promise((resolve,reject)=>{
const img = new Image();
const reader = new FileReader();
reader.onload = e=>{
img.onload = ()=>{
let {width, height} = img;
if(width>maxSize || height>maxSize){
const ratio = Math.min(maxSize/width, maxSize/height);
width = Math.round(width*ratio);
height = Math.round(height*ratio);
}
const canvas = document.createElement('canvas');
canvas.width = width; canvas.height = height;
const ctx = canvas.getContext('2d');
ctx.drawImage(img,0,0,width,height);
canvas.toBlob(blob=>{
const fr = new FileReader();
fr.onload = ()=>resolve(fr.result);
fr.onerror = reject;
fr.readAsDataURL(blob);
},'image/jpeg',quality);
};
img.onerror = reject;
img.src = e.target.result;
};
reader.onerror = reject;
reader.readAsDataURL(file);
});
}

function fmtNum(val){
const num = Number(val);
return isNaN(num)? (val||'') : num.toLocaleString();
}

// ---- Deepwell CRUD & Rendering ----
function bulletizeActivities(text){
  if(!text) return '';
  // Split by patterns like "1." "2." etc. or semicolons/double line breaks
  const parts = text.split(/\s*\d+\.\s*|\s*;\s*|\n+/).filter(p=>p.trim());
  if(parts.length<=1){
    // If splitting didn't make sense, just return original
    return text;
  }
  return `<ul class="mb-0 ps-3">${parts.map(p=>`<li>${p.trim()}</li>`).join('')}</ul>`;
}

  // ---- Deepwell CRUD & Rendering ----
function deepwellRowHtml(dw){
  let actionsHtml = (!isViewOnly && firebase.auth().currentUser) ? `<button class="btn btn-sm btn-primary me-1" title="Edit" onclick="editDeepwell('${dw.id}')"><i class="fa fa-pencil"></i></button>` : '';
  if(isAdmin){
    actionsHtml += `<button class="btn btn-sm btn-danger" title="Delete" onclick="deleteDeepwell('${dw.id}')"><i class="fa fa-trash"></i></button>`;
  }
  return `<tr data-id="${dw.id}">`
      + `<td>${dw.name}</td>`
      + `<td>${dw.provider}</td>`
      + `<td>${dw.permit||''}</td>`
      + `<td>${dw.status||''}</td>`
      + `<td>${fmtNum(dw.ratedYield)}</td>`
      + `<td>${fmtNum(dw.avgProd)}</td>`
      + `<td>${fmtNum(dw.totalProd)}</td>`
      + `<td class="d-flex gap-1">${actionsHtml}</td>`
      + `</tr>`;
}

function renderDeepwells(){
  // Show/hide Add Deepwell button based on permissions
  if(elements.addDeepwellBtn){
    elements.addDeepwellBtn.style.display = (!isViewOnly && firebase.auth().currentUser) ? 'inline-block' : 'none';
  }
  let list = deepwells;
  const provider = elements.dwProviderFilter?.value;
  const status   = elements.dwStatusFilter?.value;
  const query    = (elements.dwSearchInput?.value || '').trim().toLowerCase();
  if(provider) list = list.filter(dw=>dw.provider===provider);
  if(status)   list = list.filter(dw=>dw.status===status);
  if(query)    list = list.filter(dw => {
    const name = (dw.name||'').toLowerCase();
    const prov = (dw.provider||'').toLowerCase();
    const permit = (dw.permit||'').toLowerCase();
    return name.includes(query) || prov.includes(query) || permit.includes(query);
  });
  const tbody = elements.deepwellsTbody;
  tbody.innerHTML = list.map(deepwellRowHtml).join('');
  attachDeepwellRowEvents();
}

// ---- Deepwell Monthly Production Chart (MWCI vs MWSI) ----
function monthKeyToLabel(ym){
  if(!ym || typeof ym !== 'string' || ym.length < 7) return ym||'';
  const [y,m] = ym.split('-');
  const d = new Date(Number(y), Number(m)-1, 1);
  if(isNaN(d.getTime())) return ym;
  return d.toLocaleString('en-US', {month:'short', year:'numeric'});
}

function aggregateDeepwellMonthlyTotals(){
  const byProv = { MWCI: {}, MWSI: {} };
  (deepwells||[]).forEach(dw=>{
    const prov = (dw.provider||'').toUpperCase();
    if(!(prov in byProv)) return;
    (dw.months||[]).forEach(m=>{
      const key = m?.month;
      const val = Number(m?.prod||0);
      if(!key || !isFinite(val)) return;
      byProv[prov][key] = (byProv[prov][key]||0) + val;
    });
  });
  const months = Array.from(new Set([
    ...Object.keys(byProv.MWCI),
    ...Object.keys(byProv.MWSI)
  ])).sort();
  return {
    labels: months.map(monthKeyToLabel),
    mwci: months.map(k=>byProv.MWCI[k]||0),
    mwsi: months.map(k=>byProv.MWSI[k]||0),
    hasData: months.length>0
  };
}

function renderDeepwellMonthlyChart(){
  const canvas = elements.deepwellMonthlyChartCanvas;
  const emptyEl = elements.dwChartEmpty;
  if(!canvas) return; // not on this page
  const agg = aggregateDeepwellMonthlyTotals();
  if(!agg.hasData){
    if(emptyEl) emptyEl.style.display = 'block';
    canvas.style.display = 'none';
    if(dwMonthlyChart){ dwMonthlyChart.destroy(); dwMonthlyChart=null; }
    return;
  }
  if(emptyEl) emptyEl.style.display = 'none';
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
  const opts = {
    responsive: true,
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: { position: 'top' },
      tooltip: { callbacks: { label: (ctx)=>`${ctx.dataset.label}: ${fmtNum(ctx.parsed.y)}` } },
    },
    scales: {
      x: { title: { display: true, text: 'Month' } },
      y: { title: { display: true, text: 'Production (cu.m)' }, beginAtZero: true }
    }
  };
  if(dwMonthlyChart){
    dwMonthlyChart.data = data;
    dwMonthlyChart.options = opts;
    dwMonthlyChart.update();
  } else {
    const ctx = canvas.getContext('2d');
    dwMonthlyChart = new Chart(ctx, { type: 'line', data, options: opts });
  }
}

function attachDeepwellRowEvents(){
  const tbody = elements.deepwellsTbody;
  Array.from(tbody.querySelectorAll('tr')).forEach(row=>{
    row.addEventListener('click',e=>{
      if(isViewOnly) return;
      if(e.target.closest('button')) return;
      const id=row.dataset.id;
      editDeepwell(id,false); // open view modal (read-only)
    });
  });
}

async function saveDeepwell(dw){
  await db.collection(DEEPWELLS_COL).doc(dw.id).set(dw);
}

function gatherDwMonths(){
  const rows = Array.from(elements.dwMonthsBody.querySelectorAll('tr'));
  return rows.map(r=>({
    month: r.querySelector('.dw-month').value,
    prod: parseFloat(r.querySelector('.dw-prod').value)||0
  })).filter(m=>m.month && m.prod);
}

function updateDwStats(){
  const months = gatherDwMonths();
  const total = months.reduce((s,m)=>s+m.prod,0);
  const avg = months.length? (total/months.length):0;
  document.getElementById('dwTotalProd').value = total? total.toFixed(2):'';
  document.getElementById('dwAvgProd').value = avg? avg.toFixed(2):'';
}

elements.dwMonthsBody?.addEventListener('input',updateDwStats);

function addDwMonthRow(data={}){
  const tr=document.createElement('tr');
  tr.innerHTML=`<td><input type="month" class="form-control form-control-sm dw-month" value="${data.month||''}"/></td>
                <td><input type="number" step="0.01" class="form-control form-control-sm dw-prod" value="${data.prod||''}"/></td>
                <td class="text-end"><button type="button" class="btn btn-sm btn-outline-danger dw-remove"><i class="fa fa-minus"></i></button></td>`;
  tr.querySelector('.dw-remove').onclick=()=>{tr.remove();updateDwStats();};
  elements.dwMonthsBody.appendChild(tr);
  updateDwStats();
}

function onSaveDeepwell(e){
  e.preventDefault();
  const fd = new FormData(elements.deepwellForm);
  const id = fd.get('deepwellId') || (crypto.randomUUID?crypto.randomUUID():Date.now().toString(36));
  // Prepare history
  const userEmail = firebase.auth().currentUser?.email || 'viewer';
  const existingDw = deepwells.find(dw=>dw.id===id);
  const history = existingDw?.history ? [...existingDw.history] : [];
  history.push({email:userEmail,timestamp:new Date().toISOString(),action:existingDw?'edit':'create'});

  const deepwell = {
    id,
    name: fd.get('dwName').trim(),
    provider: fd.get('dwProvider'),
    permit: fd.get('dwPermit').trim(),
    status: fd.get('dwStatus').trim(),
    ratedYield: parseFloat(fd.get('dwRatedYield'))||0,
    months: gatherDwMonths(),
    avgProd: parseFloat(fd.get('dwAvgProd'))||0,
    totalProd: parseFloat(fd.get('dwTotalProd'))||0,
    location: fd.get('dwLocation').trim(),
    municipality: fd.get('dwMunicipality').trim(),
    history
  };
  if(!deepwell.name){alert('Deepwell Name is required');return;}
  console.log('Saving deepwell', deepwell);
saveDeepwell(deepwell)
    .then(() => {
      // Close modal and reset form
      elements.deepwellModal.hide();
      elements.deepwellForm.reset();
      elements.dwMonthsBody.innerHTML = '';
      addDwMonthRow();
      // Optimistically update local list so UI refreshes immediately
      deepwells = [...deepwells.filter(dwItem => dwItem.id !== deepwell.id), deepwell];
      renderDeepwells();
      renderDeepwellMonthlyChart();
    })
    .catch(err => alert(err.message));
}

// Attach deepwell form submit listener robustly
console.log('Attaching deepwell listeners');
const deepwellFormEl = document.getElementById('deepwellForm');
if (deepwellFormEl) {
  deepwellFormEl.addEventListener('submit', onSaveDeepwell);
}
if (elements.addDwMonthBtn) {
  elements.addDwMonthBtn.addEventListener('click', () => addDwMonthRow());
}
if (elements.addDeepwellBtn) {
  elements.addDeepwellBtn.addEventListener('click', () => {
    // Ensure modal is closed first, then reset and open
    elements.deepwellModal.hide();
    setTimeout(() => {
      resetDeepwellForm();
      elements.deepwellModal.show();
    }, 100);
  });
}

// Function to completely reset deepwell form to blank state
function resetDeepwellForm() {
  // Reset all form fields
  elements.deepwellForm.reset();
  document.getElementById('deepwellId').value = '';
  
  // Clear and reset monthly production table
  elements.dwMonthsBody.innerHTML = '';
  addDwMonthRow();
  updateDwStats();
  
  // Clear history container
  const historyContainer = document.getElementById('dwHistoryContainer');
  if (historyContainer) {
    historyContainer.innerHTML = '';
  }
  
  // Reset all form elements to editable state
  Array.from(elements.deepwellForm.elements).forEach(el => {
    if (el.tagName === 'INPUT' || el.tagName === 'SELECT') {
      el.readOnly = false;
      el.disabled = false;
      if (el.type === 'text' && ['dwRatedYield','dwAvgProd','dwTotalProd'].includes(el.id)) {
        el.type = 'number';
      }
    }
  });
  
  // Show save button
  const saveBtn = document.getElementById('saveDeepwellBtn');
  if (saveBtn) saveBtn.style.display = 'inline-block';
}

// Always clear form when the modal is fully hidden
const deepwellModalEl = document.getElementById('deepwellModal');
if (deepwellModalEl) {
  deepwellModalEl.addEventListener('hidden.bs.modal', resetDeepwellForm);
}
// Fallback: attach click listener directly to Save button in case form submit doesn't fire
const saveBtn = document.getElementById('saveDeepwellBtn');
if (saveBtn) {
  saveBtn.addEventListener('click', (e) => {
    // Show native validation UI if form is invalid
    if (!elements.deepwellForm.checkValidity()) {
      elements.deepwellForm.reportValidity();
      return;
    }
    // Prevent default submission (Bootstrap may auto-close modal)
    e.preventDefault();
    onSaveDeepwell(e);
  });
}

window.editDeepwell = (id,edit=true)=>{
  const dw = deepwells.find(x=>x.id===id);
  if(!dw) return;
  // populate form
  document.getElementById('deepwellId').value = dw.id;
  document.getElementById('dwName').value = dw.name;
  document.getElementById('dwProvider').value = dw.provider;
  document.getElementById('dwPermit').value = dw.permit||'';
  document.getElementById('dwStatus').value = dw.status||'';
  document.getElementById('dwRatedYield').value = dw.ratedYield||'';
  document.getElementById('dwAvgProd').value = dw.avgProd||'';
  document.getElementById('dwTotalProd').value = dw.totalProd||'';
  // populate months
  elements.dwMonthsBody.innerHTML='';
  (dw.months || []).forEach(m => addDwMonthRow(m));
  if ((dw.months || []).length === 0) addDwMonthRow();
  updateDwStats();
  document.getElementById('dwLocation').value = dw.location||'';
  document.getElementById('dwMunicipality').value = dw.municipality||'';
  // toggle readonly if just viewing
  // Toggle readonly/disabled state
  Array.from(elements.deepwellForm.elements).forEach(el=>{
    if(el.tagName==='INPUT' || el.tagName==='SELECT'){
      el.readOnly = !edit;
      el.disabled = !edit && el.tagName==='SELECT';
    }
  });

  // When viewing only, format number fields with commas and switch to text inputs for clarity
  const numFields = ['dwRatedYield','dwAvgProd','dwTotalProd'];
  numFields.forEach(id=>{
    const inp = document.getElementById(id);
    if(!inp) return;
    if(edit){
      // ensure numeric type while editing
      if(inp.type!=='number') inp.type='number';
    }else{
      inp.type='text';
      inp.value = fmtNum(inp.value);
    }
  });
  // Format monthly production inputs as well when viewing
  if(!edit){
    elements.dwMonthsBody.querySelectorAll('.dw-prod').forEach(inp=>{
      inp.type='text';
      inp.value = fmtNum(inp.value);
      inp.readOnly = true;
    });
    elements.dwMonthsBody.querySelectorAll('.dw-month').forEach(inp=>{inp.readOnly=true;});
  }

  // Render edit history
  const historyContainer = document.getElementById('dwHistoryContainer');
  if(historyContainer){
    if((dw.history||[]).length){
      const histSorted = dw.history.slice().sort((a,b)=> new Date(b.timestamp) - new Date(a.timestamp));
      historyContainer.innerHTML = `<h6>Edit History</h6><div class="table-responsive"><table class="table table-sm"><thead><tr><th>User</th><th>Timestamp</th><th>Action</th></tr></thead><tbody>${histSorted.map(h=>`<tr><td>${h.email}</td><td>${new Date(h.timestamp).toLocaleString()}</td><td>${h.action}</td></tr>`).join('')}</tbody></table></div>`;
    }else{
      historyContainer.innerHTML = '';
    }
  }

  document.getElementById('saveDeepwellBtn').style.display = edit?'inline-block':'none';
  elements.deepwellModal.show();
};

window.deleteDeepwell = async id=>{
  if(!elevatedAccess){alert('Only admin/level2 can delete deepwells');return;}
  if(!confirm('Delete this deepwell?')) return;
  try{
    await db.collection(DEEPWELLS_COL).doc(id).delete();
  }catch(err){alert(err.message);} 
};

// Tab switching
function showProjectsSection(){
  elements.projectsTab.classList.add('active');
  elements.deepwellsTab.classList.remove('active');
  elements.projectsSection.style.display='block';
  elements.deepwellsSection.style.display='none';
  elements.reforestationSection.style.display='none';
}
function showDeepwellsSection(){
  elements.projectsTab.classList.remove('active');
  elements.deepwellsTab.classList.add('active');
  elements.projectsSection.style.display='none';
  elements.deepwellsSection.style.display='block';
  elements.reforestationSection.style.display='none';
  renderDeepwells();
  renderDeepwellMonthlyChart();
}

elements.projectsTab?.addEventListener('click',e=>{e.preventDefault();showProjectsSection();});
elements.deepwellsTab?.addEventListener('click',e=>{e.preventDefault();showDeepwellsSection();});

function showReforestationSection(){
  elements.projectsTab.classList.remove('active');
  elements.deepwellsTab.classList.remove('active');
  elements.reforestationTab.classList.add('active');
  elements.projectsSection.style.display='none';
  elements.deepwellsSection.style.display='none';
  elements.reforestationSection.style.display='block';
  renderReforestations();
}

elements.reforestationTab?.addEventListener('click',e=>{e.preventDefault();showReforestationSection();});

// Add Reforestation button opens modal with cleared form
if(elements.addReforestationBtn){
  elements.addReforestationBtn.addEventListener('click',()=>{
    elements.reforestationForm.reset();
    elements.reforestationForm.reforestationId.value='';
    document.getElementById('saveReforestationBtn').style.display='inline-block';
    elements.reforestationModal.show();
  });
}


// Subscribe to deepwells data when authenticated
function subscribeDeepwells(){
  if(unsubscribeDeepwells) unsubscribeDeepwells();
  unsubscribeDeepwells = db.collection(DEEPWELLS_COL).onSnapshot(snap=>{
    deepwells = snap.docs.map(d=>({id:d.id,...d.data()}));
    renderDeepwells();
    renderDeepwellMonthlyChart();
  });
}

// integrate into auth listener: call subscribeDeepwells() when user signed in / anon / viewer
// called after projects subscription below

// ---- Reforestation CRUD & Rendering ----
function reforestationRowHtml(r){
  const deleteBtnHtml = isAdmin ? `<button class='btn btn-sm btn-danger del-ref'><i class='fa fa-trash'></i></button>` : '';
  const editBtn = `<button class='btn btn-sm btn-primary edit-ref'><i class='fa fa-pen'></i></button>`;
  return `<tr data-id="${r.id}">
    <td>${r.activityName||''}</td>
    <td>${r.activityType||''}</td>
    <td>${r.location||''}</td>
    <td>${r.activityStatus||''}</td>
    <td>${r.targetDate||''}</td>
    <td>${fmtNum(r.treesPlanted||0)}</td>
    <td class='d-flex gap-1'>${editBtn}${deleteBtnHtml}</td>
  </tr>`;
}
function renderReforestations(){
  if(!elements.reforestationTbody) return;
  const text   = (elements.reforestationSearchInput?.value || '').toLowerCase();
  const rType  = elements.reforestationTypeFilter?.value || '';
  const rStatus= elements.reforestationStatusFilter?.value || '';
  const filtered = reforestations.filter(r=>{
    const matchText   = (r.activityName||'').toLowerCase().includes(text) || (r.location||'').toLowerCase().includes(text);
    const matchType   = rType ? r.activityType === rType : true;
    const matchStatus = rStatus ? r.activityStatus === rStatus : true;
    return matchText && matchType && matchStatus;
  });
  elements.reforestationTbody.innerHTML = filtered.map(reforestationRowHtml).join('');
  attachReforestationRowEvents();
}
function attachReforestationRowEvents(){
  // Make whole row clickable (excluding action buttons) for viewing details
  document.querySelectorAll('#reforestationTable tbody tr').forEach(tr=>{
    tr.style.cursor = 'pointer';
    tr.onclick = e=>{
      // Ignore clicks on the edit/delete buttons
      if(e.target.closest('.edit-ref') || e.target.closest('.del-ref')) return;
      viewReforestation(tr.dataset.id);
    };
  });
  // Existing edit/delete buttons
    document.querySelectorAll('#reforestationTable .edit-ref').forEach(btn=>{
    btn.onclick = ()=>{
      const id = btn.closest('tr').dataset.id;
      editReforestation(id);
    };
  });
  document.querySelectorAll('#reforestationTable .del-ref').forEach(btn=>{
    btn.onclick = ()=>{
      const id = btn.closest('tr').dataset.id;
      deleteReforestation(id);
    };
  });
}

// Show Reforestation Activity details (view only)
window.viewReforestation = (id)=>{
  const r = reforestations.find(x=>x.id===id);
  if(!r) return;
  const body = elements.reforestationDetailsBody;
   const histSorted = (r.history||[]).slice().sort((a,b)=> new Date(b.timestamp) - new Date(a.timestamp));
  body.innerHTML = `
    <h5 class="fw-bold mb-2">${r.activityName||''}</h5>
    <p><strong>Type:</strong> ${r.activityType||''}</p>
    <p><strong>Location:</strong> ${r.location||''}</p>
    <p><strong>Implementing Agency:</strong> ${r.implementingAgency||''}</p>
    <p><strong>Status:</strong> ${r.activityStatus||''}</p>
    <p><strong>Start Date:</strong> ${r.startDate||''}</p>
    <p><strong>Target Date:</strong> ${r.targetDate||''}</p>
    <p><strong>Target Area (ha):</strong> ${r.targetArea||0}</p>
    <p><strong>Trees Planted:</strong> ${fmtNum(r.treesPlanted||0)}</p>
    <p><strong>Tree Species:</strong> ${r.treeSpecies||''}</p>
    <p><strong>Budget (PHP):</strong> ₱${Number(r.budget||0).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</p>
    <p><strong>Initial Survival Rate:</strong> ${r.initialSurvivalRate||0}% ${r.initialSurvivalDate?`(as of ${r.initialSurvivalDate})`:''}</p>
    <p><strong>Final Survival Rate:</strong> ${r.finalSurvivalRate||0}% ${r.finalSurvivalDate?`(as of ${r.finalSurvivalDate})`:''}</p>
    ${r.description?`<p><strong>Description:</strong> ${r.description}</p>`:''}
    ${r.kmzName?`<p><strong>KMZ:</strong> <a href="${r.kmzDataUrl}" download="${r.kmzName}">${r.kmzName}</a></p>`:''}
    ${r.remarksReforestation?`<p><strong>Remarks:</strong> ${r.remarksReforestation}</p>`:''}
    ${r.photos&&r.photos.length?`<div class="d-flex gap-3 flex-wrap mt-3">${r.photos.map(url=>`<img src="${url}" style="max-height:250px;object-fit:contain;" class="img-fluid border">`).join('')}</div>`:''}
    ${histSorted.length?`<h6 class="mt-4">Edit History</h6><div class="table-responsive"><table class="table table-sm"><thead><tr><th>User</th><th>Timestamp</th><th>Action</th></tr></thead><tbody>${histSorted.map(h=>`<tr><td>${h.email}</td><td>${new Date(h.timestamp).toLocaleString()}</td><td>${h.action}</td></tr>`).join('')}</tbody></table></div>`:''}
  `;
  elements.reforestationDetailsModal.show();
};

async function saveReforestation(r){
  let existingKmzName = '';
  let existingKmzDataUrl = '';
  // Fetch existing doc if editing to preserve current photos
  let existingPhotos = [];
  if(r.id){
    try{
      const docSnap = await db.collection(REFORESTATION_COL).doc(r.id).get();
      if(docSnap.exists){
        const docData = docSnap.data();
        existingPhotos = docData.photos || [];
        existingKmzName = docData.kmzName || '';
        existingKmzDataUrl = docData.kmzDataUrl || '';
      }
    }catch(e){ console.error('Failed to fetch existing reforestation doc',e); }
  }
  
  // Handle photo uploads (max 3)
  const photoInputs=["reforestPhoto1","reforestPhoto2","reforestPhoto3"].map(id=>document.getElementById(id));
  const files=photoInputs.map(inp=>inp?.files[0]).filter(f=>!!f);
  let newPhotos = [];
  // --- KMZ processing vars ---
  const kmzInput = document.getElementById('kmzFile');
  const kmzFile  = kmzInput?.files[0] || null;
  let kmzName = existingKmzName || '';
  let kmzDataUrl = existingKmzDataUrl || '';
  if(files.length){
    try{
      newPhotos = await Promise.all(files.map(f=>compressImage(f,1024,0.75)));
    }catch(err){alert('Error reading photos: '+err.message);}  }
  // Final photos array – use new uploads if any, else keep existing
  // Persist photos (replace only if new ones uploaded)
  r.photos = newPhotos.length ? newPhotos.slice(0,3) : existingPhotos;
  // --- KMZ upload logic ---
  if(kmzFile){
    if(!kmzFile.name.toLowerCase().endsWith('.kmz')){
      alert('Only .kmz files are allowed');
      return;
    }
    kmzName = kmzFile.name;
    // Read binary as Data URL so we can store in Firestore safely (<1 MB recommended)
    try{
      kmzDataUrl = await new Promise((res,rej)=>{
        const reader = new FileReader();
        reader.onload = ()=>res(reader.result);
        reader.onerror = rej;
        reader.readAsDataURL(kmzFile);
      });
    }catch(err){ alert('Error reading KMZ file: '+err.message); return; }
  }
  r.kmzName = kmzName;
  r.kmzDataUrl = kmzDataUrl;
  // --- Edit History ---
  const userEmail = firebase.auth().currentUser?.email || 'unknown';
  const existingHistory = r.id ? (await db.collection(REFORESTATION_COL).doc(r.id).get()).data().history || [] : [];
  const actionLabel = r.id ? 'edit' : 'create';
  r.history = [...existingHistory, {email:userEmail,timestamp:new Date().toISOString(),action:actionLabel}];

  const data = {...r};
  if(!data.id) delete data.id; // Firestore cannot store undefined
  if(r.id){
    await db.collection(REFORESTATION_COL).doc(r.id).set(data);
  }else{
    await db.collection(REFORESTATION_COL).add(data);
  }
}
function clearReforestationForm(){
  const f = elements.reforestationForm;
  f.reset();
  if(f.reforestationId) f.reforestationId.value='';
  ["reforestPhoto1","reforestPhoto2","reforestPhoto3"].forEach(id=>{const el=document.getElementById(id);if(el) el.value="";});
  const kmzInp=document.getElementById('kmzFile'); if(kmzInp) kmzInp.value="";
}

function gatherReforestationForm(){
  const f = elements.reforestationForm;
  return {
    id: f.reforestationId.value || undefined,
    activityName: f.activityName.value,
    activityType: f.activityType.value,
    location: f.location.value,
    implementingAgency: f.implementingAgency.value,
    targetArea: parseFloat(f.targetArea.value) || 0,
    treesPlanted: parseInt(f.treesPlanted.value) || 0,
    treeSpecies: f.treeSpecies.value,
    activityStatus: f.activityStatus.value,
    startDate: f.startDate.value,
    targetDate: f.targetDate.value,
    budget: parseFloat(f.budget.value) || 0,
    // --- new survival fields ---
    initialSurvivalRate: parseFloat(f.initialSurvivalRate.value) || 0,
    initialSurvivalDate: f.initialSurvivalDate.value,
    finalSurvivalRate: parseFloat(f.finalSurvivalRate.value) || 0,
    finalSurvivalDate: f.finalSurvivalDate.value,
    // ---------------------------------
    description: f.description.value,
    remarksReforestation: f.remarksReforestation.value,
    timestamp: Date.now()
  };
}
function onSaveReforestation(e){
  e.preventDefault();
  const saveBtn = document.getElementById('saveReforestationBtn');
  saveBtn.disabled = true;
  const data = gatherReforestationForm();
  // Close modal immediately for better UX
  elements.reforestationModal.hide();
  saveReforestation(data)
    .catch(err=>{alert(err.message);})
    .finally(()=>{saveBtn.disabled=false;});
}
function editReforestation(id){
  const r = reforestations.find(r=>r.id===id);
  if(!r) return;
  const f = elements.reforestationForm;
  Object.keys(r).forEach(k=>{ if(f[k]) f[k].value = r[k];});
  // set hidden id field
  if(f.reforestationId) f.reforestationId.value = r.id;
  document.getElementById('saveReforestationBtn').style.display='inline-block';
  elements.reforestationModal.show();
}
async function deleteReforestation(id){
  if(!isAdmin){alert('Only admin can delete');return;}
  if(!confirm('Delete this activity?')) return;
  await db.collection(REFORESTATION_COL).doc(id).delete();
}

// Attach listener
if(elements.reforestationForm){ elements.reforestationForm.addEventListener('submit',onSaveReforestation);} 

function subscribeReforestations(){
  if(unsubscribeReforestations) unsubscribeReforestations();
  unsubscribeReforestations = db.collection(REFORESTATION_COL).onSnapshot(snap=>{
    reforestations = snap.docs.map(d=>({id:d.id,...d.data()}));
    renderReforestations();
  });
}

// ---- End project functions ----

// Signup handler (guard if signup form absent on mobile)
  if(signupForm){
    signupForm.addEventListener('submit', e=>{
      e.preventDefault();
      const email=document.getElementById('signupEmail').value.trim();
      const pw1=document.getElementById('signupPassword').value;
      const pw2=document.getElementById('signupPassword2').value;
      if(pw1!==pw2){alert('Passwords do not match');return;}
      firebase.auth().createUserWithEmailAndPassword(email,pw1)
        .then(cred=>{
          // mark as unapproved
          return db.collection('users').doc(cred.user.uid).set({
            email: cred.user.email,
            approved: false,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
          }).then(()=>{
            alert('Account request submitted. An admin will review your request.');
            firebase.auth().signOut();
            showLoginForm();
          });
        })
        .catch(err=>alert(err.message));
    });
  }

  
  

function showLogin() {
loginScreen.classList.remove('d-none');
loginScreen.style.display = 'flex';
appWrapper.style.display  = 'none';
logoutBtn.style.display   = 'none';
}

function showSignup(){
  if(loginForm) loginForm.classList.add('d-none');
  if(loginLinks) loginLinks.classList.add('d-none');
  if(signupForm) signupForm.classList.remove('d-none');
  if(signupLinks) signupLinks.classList.remove('d-none');
  const loginTitle = document.getElementById('loginTitle');
  if(loginTitle) loginTitle.textContent='Sign Up';
}

function showLoginForm(){
  if(signupForm) signupForm.classList.add('d-none');
  if(signupLinks) signupLinks.classList.add('d-none');
  if(loginForm) loginForm.classList.remove('d-none');
  if(loginLinks) loginLinks.classList.remove('d-none');
  const loginTitle = document.getElementById('loginTitle');
  if(loginTitle) loginTitle.textContent='Login';
}

// link listeners
if(showSignupLink) showSignupLink.addEventListener('click',e=>{e.preventDefault();showSignup();});
if(showLoginLink)  showLoginLink.addEventListener('click',e=>{e.preventDefault();showLoginForm();});

// Forgot password
const forgotPwLink = document.getElementById('forgotPwLink');
if(forgotPwLink){
  forgotPwLink.addEventListener('click', async e => {
    e.preventDefault();
    // Try to pre-fill with whatever is typed in the email field, if available
    const loginEmailInput = document.querySelector('#loginForm input[type="email"]');
    const defaultEmail = loginEmailInput ? loginEmailInput.value.trim() : '';
    const email = prompt('Enter your registered email address:', defaultEmail);
    if(!email) return;
    try{
      await firebase.auth().sendPasswordResetEmail(email.trim());
      alert('Password-reset email sent. Please check your inbox (and spam folder).');
    }catch(err){
      alert('Failed to send reset email: '+err.message);
    }
  });
}

function showApp() {
  if(isViewOnly){
    loginScreen.classList.add('d-none');
    appWrapper.style.display='block';
    loginBtn.style.display = 'inline-block';
    logoutBtn.style.display='none';
    return;
  }
  if(isViewOnly){
    loginScreen.classList.add('d-none');
    appWrapper.style.display='block';
    logoutBtn.style.display='none';
    return;
  }
loginScreen.classList.add('d-none');
appWrapper.style.display  = 'block';
loginBtn.style.display   = 'none';
logoutBtn.style.display   = 'inline-block';
}


  // Monitor auth state
  firebase.auth().onAuthStateChanged(async user => {
    if (user) {
      if(user.isAnonymous){
        isViewOnly = true;
        isAdmin = false;
        showApp();
        updateAdminUI();
      subscribeDeepwells();
        if(unsubscribeProjects) unsubscribeProjects();
        unsubscribeProjects = db.collection(PROJECTS_COL).onSnapshot(snap=>{
          projects = snap.docs.map(d=>({id:d.id,...d.data()}));
          render();
        });
        return;
      }
      const emailLower = (user.email||'').trim().toLowerCase();
      isAdmin = emailLower === ADMIN_EMAIL_LOWER;
      console.log('Signed-in as', user.email, '| isAdmin?', isAdmin);
      // Determine user access level
      if(isAdmin){
        isLevel2 = false;
        elevatedAccess = true;
      }else{
        const doc = await db.collection('users').doc(user.uid).get();
        if(!doc.exists || doc.data().approved!==true){
          alert('Your account is pending approval.');
          firebase.auth().signOut();
          return;
        }
        const lvl = doc.data().accessLevel || 1;
        isLevel2 = lvl === 2;
        elevatedAccess = isLevel2; // admins already handled above
      }
      showApp();
      updateAdminUI();
      // Ensure approved users list is available for chat recipients even for non-admins
      if(!isAdmin){
        loadApprovedUsers();
      }
      subscribeDeepwells();
      if(unsubscribeProjects) unsubscribeProjects();
      const legacyKey='construction_projects';
async function migrateLegacyIfAny(){
  try{
    const legacy = JSON.parse(localStorage.getItem(legacyKey)||'[]');
    if(!legacy.length) return false;
    const batch = db.batch();
    legacy.forEach(p=>{
      const id=p.id|| (crypto.randomUUID?crypto.randomUUID():Date.now().toString(36));
      batch.set(db.collection(PROJECTS_COL).doc(id), {...p,id});
    });
    await batch.commit();
    localStorage.removeItem(legacyKey);
    console.log('Migrated',legacy.length,'legacy projects to Firestore');
    return true;
  }catch(err){console.error('Legacy migration failed',err); return false;}
}

await migrateLegacyIfAny();

unsubscribeProjects = db.collection(PROJECTS_COL).onSnapshot(async snap => {
        projects = snap.docs.map(d => ({ id: d.id, ...d.data() }));
        if(projects.length===0){
          // attempt legacy localStorage migration
          try{
            const legacy = JSON.parse(localStorage.getItem('construction_projects')||'[]');
            if(legacy.length){
              const batch = db.batch();
              legacy.forEach(p=>{
                const id = p.id || (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36));
                batch.set(db.collection(PROJECTS_COL).doc(id), {...p,id});
              });
              await batch.commit();
              localStorage.removeItem('construction_projects');
              console.log('Migrated', legacy.length,'projects from localStorage to Firestore');
            }
          }catch(err){console.warn('Legacy migration failed',err);}
        }
        render();
      });
    } else {
      // No user signed in: stay on the login screen until the user authenticates.
      isAdmin = false;
      isViewOnly = false;
      showLogin();
      updateAdminUI();
      // Clean up any active listeners from previous sessions
      if(unsubscribeProjects){unsubscribeProjects(); unsubscribeProjects=null;}
      if(unsubscribeDeepwells){unsubscribeDeepwells(); unsubscribeDeepwells=null;}
      projects = [];
      deepwells = [];
      render();
    }
  });

  // Handle login form submission (guard for mobile if login form missing)
  if(loginForm){
    loginForm.addEventListener('submit', e => {
      e.preventDefault();
      const email = document.getElementById('loginEmail').value.trim();
      const pw    = document.getElementById('loginPassword').value;
      firebase.auth().signInWithEmailAndPassword(email, pw)
        .catch(err => alert(err.message));
    });
  }

  // Handle logout (guard)
  if(logoutBtn){
    logoutBtn.addEventListener('click', () => firebase.auth().signOut());
  }
  if(loginBtn) loginBtn.addEventListener('click', () => {
    // exit view-only: go back to login page
    isViewOnly = false;
    firebase.auth().signOut().catch(()=>{}); // ensure no anon session remains
    showLogin();
  });
  // Admin manage pending users
  elements.manageUsersBtn.addEventListener('click', ()=>{
    if(!isAdmin) return;
    loadPendingUsers();
    loadApprovedUsers();
    bootstrap.Modal.getOrCreateInstance(document.getElementById('usersModal')).show();
  });


})();
