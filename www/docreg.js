// Document Registry (Outgoing Letter) - Standalone page logic
// Data is stored in localStorage for now. Replace with Firestore later if needed.

(function () {
  const els = {
    // section toggles
    btnOutgoing: document.getElementById('btnOutgoing'),
    btnIncoming: document.getElementById('btnIncoming'),
    btnMemo: document.getElementById('btnMemo'),
    btnOffice: document.getElementById('btnOffice'),
    btnTravel: document.getElementById('btnTravel'),
    btnMinutes: document.getElementById('btnMinutes'),
    btnNotice: document.getElementById('btnNotice'),
    outgoingSection: document.getElementById('outgoingSection'),
    incomingSection: document.getElementById('incomingSection'),
    memoSection: document.getElementById('memoSection'),
    officeSection: document.getElementById('officeSection'),
    travelSection: document.getElementById('travelSection'),
    minutesSection: document.getElementById('minutesSection'),
    noticeSection: document.getElementById('noticeSection'),
    addBtn: document.getElementById('olAddBtn'),
    tbody: document.getElementById('olTbody'),
    empty: document.getElementById('olEmpty'),
    form: document.getElementById('olForm'),
    modalEl: document.getElementById('olModal'),
    detailsModalEl: document.getElementById('olDetailsModal'),
    detailsBody: document.getElementById('olDetailsBody'),
    // action bar controls
    searchInput: document.getElementById('olSearchInput'),
    statusFilter: document.getElementById('olStatusFilter'),
    deliveryFilter: document.getElementById('olDeliveryFilter'),
    refNo: document.getElementById('olRefNo'),
    docDate: document.getElementById('olDocDate'),
    controlNo: document.getElementById('olControlNo'),
    olPartyBlock: document.getElementById('olPartyBlock'),
    name: document.getElementById('olName'),
    designation: document.getElementById('olDesignation'),
    agency: document.getElementById('olAgency'),
    address: document.getElementById('olAddress'),
    // New structured TO/FROM for Outgoing
    olFromName: document.getElementById('olFromName'),
    olFromDesignation: document.getElementById('olFromDesignation'),
    olFromAgency: document.getElementById('olFromAgency'),
    olFromAddress: document.getElementById('olFromAddress'),
    olToName: document.getElementById('olToName'),
    olToPosition: document.getElementById('olToPosition'),
    olToAgency: document.getElementById('olToAgency'),
    olToAddress: document.getElementById('olToAddress'),
    subject: document.getElementById('olSubject'),
    docLink: document.getElementById('olDocLink'),
    preparedBy: document.getElementById('olPreparedBy'),
    deliveryVia: document.getElementById('olDeliveryVia'),
    status: document.getElementById('olStatus'),
    remarks: document.getElementById('olRemarks'),
    signatory: document.getElementById('olSignatory'),
    id: document.getElementById('olId'),
    // Incoming elements
    ilAddBtn: document.getElementById('ilAddBtn'),
    ilTbody: document.getElementById('ilTbody'),
    ilEmpty: document.getElementById('ilEmpty'),
    ilForm: document.getElementById('ilForm'),
    ilModalEl: document.getElementById('ilModal'),
    ilDetailsModalEl: document.getElementById('ilDetailsModal'),
    ilDetailsBody: document.getElementById('ilDetailsBody'),
    ilSearchInput: document.getElementById('ilSearchInput'),
    ilStatusFilter: document.getElementById('ilStatusFilter'),
    ilDeliveryFilter: document.getElementById('ilDeliveryFilter'),
    ilRefNo: document.getElementById('ilRefNo'),
    ilDocDate: document.getElementById('ilDocDate'),
    // Legacy single block (kept for back-compat)
    ilName: document.getElementById('ilName'),
    ilDesignation: document.getElementById('ilDesignation'),
    ilAgency: document.getElementById('ilAgency'),
    ilAddress: document.getElementById('ilAddress'),
    // New structured TO/FROM fields
    ilToName: document.getElementById('ilToName'),
    ilToDesignation: document.getElementById('ilToDesignation'),
    ilToAgency: document.getElementById('ilToAgency'),
    ilToAddress: document.getElementById('ilToAddress'),
    ilFromName: document.getElementById('ilFromName'),
    ilFromPosition: document.getElementById('ilFromPosition'),
    ilFromAgency: document.getElementById('ilFromAgency'),
    ilFromAddress: document.getElementById('ilFromAddress'),
    ilPartyBlock: document.getElementById('ilPartyBlock'),
    ilSubject: document.getElementById('ilSubject'),
    ilDocLink: document.getElementById('ilDocLink'),
    ilCategory: document.getElementById('ilCategory'),
    ilReceivedDate: document.getElementById('ilReceivedDate'),
    ilDeliveryVia: document.getElementById('ilDeliveryVia'),
    ilAssignedTo: document.getElementById('ilAssignedTo'),
    ilStatus: document.getElementById('ilStatus'),
    ilRemarks: document.getElementById('ilRemarks'),
    ilDeliveryDate: document.getElementById('ilDeliveryDate'),
    ilId: document.getElementById('ilId'),
    // Memorandum elements
    moAddBtn: document.getElementById('moAddBtn'),
    moTbody: document.getElementById('moTbody'),
    moEmpty: document.getElementById('moEmpty'),
    moForm: document.getElementById('moForm'),
    moModalEl: document.getElementById('moModal'),
    moDetailsModalEl: document.getElementById('moDetailsModal'),
    moDetailsBody: document.getElementById('moDetailsBody'),
    moSearchInput: document.getElementById('moSearchInput'),
    moStatusFilter: document.getElementById('moStatusFilter'),
    moRefNo: document.getElementById('moRefNo'),
    moControlNo: document.getElementById('moControlNo'),
    moDocDate: document.getElementById('moDocDate'),
    moPartyBlock: document.getElementById('moPartyBlock'),
    // Memorandum TO/FROM structured fields (new)
    moToName: document.getElementById('moToName'),
    moToDesignation: document.getElementById('moToDesignation'),
    moToDepartment: document.getElementById('moToDepartment'),
    moFromName: document.getElementById('moFromName'),
    moFromDesignation: document.getElementById('moFromDesignation'),
    moFromDepartment: document.getElementById('moFromDepartment'),
    moSubject: document.getElementById('moSubject'),
    moDocLink: document.getElementById('moDocLink'),
    moPreparedBy: document.getElementById('moPreparedBy'),
    moSignatory: document.getElementById('moSignatory'),
    moStatus: document.getElementById('moStatus'),
    moRemarks: document.getElementById('moRemarks'),
    moId: document.getElementById('moId'),
    // Office Order elements
    ooAddBtn: document.getElementById('ooAddBtn'),
    ooTbody: document.getElementById('ooTbody'),
    ooEmpty: document.getElementById('ooEmpty'),
    ooForm: document.getElementById('ooForm'),
    ooModalEl: document.getElementById('ooModal'),
    ooDetailsModalEl: document.getElementById('ooDetailsModal'),
    ooDetailsBody: document.getElementById('ooDetailsBody'),
    ooSearchInput: document.getElementById('ooSearchInput'),
    ooRefNo: document.getElementById('ooRefNo'),
    ooDocDate: document.getElementById('ooDocDate'),
    ooIncStart: document.getElementById('ooIncStart'),
    ooIncEnd: document.getElementById('ooIncEnd'),
    ooNoEnd: document.getElementById('ooNoEnd'),
    ooModeRange: document.getElementById('ooModeRange'),
    ooModeSingle: document.getElementById('ooModeSingle'),
    ooModeMulti: document.getElementById('ooModeMulti'),
    ooRangeWrap: document.getElementById('ooRangeWrap'),
    ooSingleWrap: document.getElementById('ooSingleWrap'),
    ooMultiWrap: document.getElementById('ooMultiWrap'),
    ooSingleDate: document.getElementById('ooSingleDate'),
    ooMultiDateInput: document.getElementById('ooMultiDateInput'),
    ooMultiAddBtn: document.getElementById('ooMultiAddBtn'),
    ooMultiList: document.getElementById('ooMultiList'),
    ooSubject: document.getElementById('ooSubject'),
    ooParticipants: document.getElementById('ooParticipants'),
    ooDocLink: document.getElementById('ooDocLink'),
    ooId: document.getElementById('ooId'),
    // Travel Order elements
    toAddBtn: document.getElementById('toAddBtn'),
    toTbody: document.getElementById('toTbody'),
    toEmpty: document.getElementById('toEmpty'),
    toForm: document.getElementById('toForm'),
    toModalEl: document.getElementById('toModal'),
    toDetailsModalEl: document.getElementById('toDetailsModal'),
    toDetailsBody: document.getElementById('toDetailsBody'),
    toSearchInput: document.getElementById('toSearchInput'),
    toRefNo: document.getElementById('toRefNo'),
    toDocDate: document.getElementById('toDocDate'),
    toIncStart: document.getElementById('toIncStart'),
    toIncEnd: document.getElementById('toIncEnd'),
    toNoEnd: document.getElementById('toNoEnd'),
    toModeRange: document.getElementById('toModeRange'),
    toModeSingle: document.getElementById('toModeSingle'),
    toModeMulti: document.getElementById('toModeMulti'),
    toRangeWrap: document.getElementById('toRangeWrap'),
    toSingleWrap: document.getElementById('toSingleWrap'),
    toMultiWrap: document.getElementById('toMultiWrap'),
    toSingleDate: document.getElementById('toSingleDate'),
    toMultiDateInput: document.getElementById('toMultiDateInput'),
    toMultiAddBtn: document.getElementById('toMultiAddBtn'),
    toMultiList: document.getElementById('toMultiList'),
    toSubject: document.getElementById('toSubject'),
    toParticipants: document.getElementById('toParticipants'),
    toDocLink: document.getElementById('toDocLink'),
    toId: document.getElementById('toId'),
    // Minutes of Meeting elements
    miAddBtn: document.getElementById('miAddBtn'),
    miTbody: document.getElementById('miTbody'),
    miEmpty: document.getElementById('miEmpty'),
    miForm: document.getElementById('miForm'),
    miModalEl: document.getElementById('miModal'),
    miDetailsModalEl: document.getElementById('miDetailsModal'),
    miDetailsBody: document.getElementById('miDetailsBody'),
    miSearchInput: document.getElementById('miSearchInput'),
    miControlNo: document.getElementById('miControlNo'),
    miDocDate: document.getElementById('miDocDate'),
    miMeetDate: document.getElementById('miMeetDate'),
    miSignatory: document.getElementById('miSignatory'),
    miSubject: document.getElementById('miSubject'),
    miAgenda: document.getElementById('miAgenda'),
    miPreparedBy: document.getElementById('miPreparedBy'),
    miDocLink: document.getElementById('miDocLink'),
    miRemarks: document.getElementById('miRemarks'),
    miId: document.getElementById('miId'),
    // Notice of Meeting elements
    noAddBtn: document.getElementById('noAddBtn'),
    noTbody: document.getElementById('noTbody'),
    noEmpty: document.getElementById('noEmpty'),
    noForm: document.getElementById('noForm'),
    noModalEl: document.getElementById('noModal'),
    noDetailsModalEl: document.getElementById('noDetailsModal'),
    noDetailsBody: document.getElementById('noDetailsBody'),
    noSearchInput: document.getElementById('noSearchInput'),
    noRefNo: document.getElementById('noRefNo'),
    noControlNo: document.getElementById('noControlNo'),
    noDocDate: document.getElementById('noDocDate'),
    noMeetDate: document.getElementById('noMeetDate'),
    noMeetTime: document.getElementById('noMeetTime'),
    noPartyBlock: document.getElementById('noPartyBlock'),
    noSubject: document.getElementById('noSubject'),
    noSignatory: document.getElementById('noSignatory'),
    noSignatoryBlock: document.getElementById('noSignatoryBlock'),
    noSignatoryDesignation: document.getElementById('noSignatoryDesignation'),
    noVenue: document.getElementById('noVenue'),
    noDocLink: document.getElementById('noDocLink'),
    noRemarks: document.getElementById('noRemarks'),
    noId: document.getElementById('noId'),
    // Spinner
    loadingSpinner: document.getElementById('loadingSpinner'),
  };

  // Global loading spinner helper
  window.showLoading = (show) => {
    if (els.loadingSpinner) {
      if (show) els.loadingSpinner.classList.add('show');
      else els.loadingSpinner.classList.remove('show');
    }
  };

  // Continue even if some elements are missing; guards are applied below so that
  // the page remains functional and event listeners attach where possible.

  // Ensure Incoming modals are not nested inside the sticky header (which can
  // create stacking-context issues that block typing). If found under the
  // header, move them to <body> before initializing Bootstrap.Modal instances.
  try {
    const header = document.getElementById('logosHeader');
    if (header) {
      if (els.ilModalEl && header.contains(els.ilModalEl)) {
        document.body.appendChild(els.ilModalEl);
      }
      function ilUnlinkReply(incomingId, replyId) {
        // 1) Remove from incoming.replies (local and Firestore)
        try {
          const items = ilLoad();
          const idx = items.findIndex(x => x.id === incomingId);
          if (idx >= 0) {
            const r = (items[idx].replies || []).filter(x => x && x.id !== replyId);
            const unlinkedSet = new Set(Array.isArray(items[idx].unlinkedReplyIds) ? items[idx].unlinkedReplyIds : []);
            unlinkedSet.add(replyId);
            const unlinkedArr = Array.from(unlinkedSet);
            items[idx] = { ...items[idx], replies: r, unlinkedReplyIds: unlinkedArr };
            ilSave(items);
            try {
              const db = window.firebase?.firestore?.();
              if (db) {
                const FieldValue = window.firebase?.firestore?.FieldValue;
                const payload = FieldValue
                  ? { replies: r, unlinkedReplyIds: FieldValue.arrayUnion(replyId) }
                  : { replies: r, unlinkedReplyIds: unlinkedArr };
                db.collection('docreg_incoming').doc(incomingId).set(payload, { merge: true });
              }
            } catch (_) { }
          }
        } catch (_) { }
        // 2) Clear explicit linkage on Outgoing (local)
        try {
          const outs = loadItems();
          const oIdx = outs.findIndex(x => x.id === replyId);
          if (oIdx >= 0 && outs[oIdx].replyToIncomingId === incomingId) { const obj = { ...outs[oIdx] }; delete obj.replyToIncomingId; outs[oIdx] = obj; saveItems(outs); }
        } catch (_) { }
        try {
          const mem = moLoad();
          const mIdx = mem.findIndex(x => x.id === replyId);
          if (mIdx >= 0 && mem[mIdx].replyToIncomingId === incomingId) { const obj = { ...mem[mIdx] }; delete obj.replyToIncomingId; mem[mIdx] = obj; moSave(mem); }
        } catch (_) { }
        // 5) Refresh details
        try { ilOpenDetails(incomingId); } catch (_) { }
      }

      function ilRemoveUnlinked(incomingId, replyId) {
        if (!incomingId || !replyId) return;
        try {
          const items = ilLoad();
          const idx = items.findIndex(x => x.id === incomingId);
          if (idx >= 0) {
            const arr = Array.isArray(items[idx].unlinkedReplyIds) ? items[idx].unlinkedReplyIds.filter(id => id !== replyId) : [];
            items[idx] = { ...items[idx], unlinkedReplyIds: arr };
            ilSave(items);
          }
        } catch (_) { }
        try {
          const FieldValue = window.firebase?.firestore?.FieldValue;
          const db = window.firebase?.firestore?.();
          if (db && FieldValue) {
            db.collection('docreg_incoming').doc(incomingId).set({ unlinkedReplyIds: FieldValue.arrayRemove(replyId) }, { merge: true });
          }
        } catch (_) { }
      }
      if (els.ilDetailsModalEl && header.contains(els.ilDetailsModalEl)) {
        document.body.appendChild(els.ilDetailsModalEl);
      }
      const replyChoiceEl = document.getElementById('replyChoiceModal');
      if (replyChoiceEl && header.contains(replyChoiceEl)) {
        document.body.appendChild(replyChoiceEl);
      }
      if (els.moModalEl && header.contains(els.moModalEl)) {
        document.body.appendChild(els.moModalEl);
      }
      if (els.moDetailsModalEl && header.contains(els.moDetailsModalEl)) {
        document.body.appendChild(els.moDetailsModalEl);
      }
      if (els.ooModalEl && header.contains(els.ooModalEl)) {
        document.body.appendChild(els.ooModalEl);
      }
      if (els.ooDetailsModalEl && header.contains(els.ooDetailsModalEl)) {
        document.body.appendChild(els.ooDetailsModalEl);
      }
      // Travel Order modals
      if (els.toModalEl && header.contains(els.toModalEl)) {
        document.body.appendChild(els.toModalEl);
      }
      if (els.toDetailsModalEl && header.contains(els.toDetailsModalEl)) {
        document.body.appendChild(els.toDetailsModalEl);
      }
      // Minutes of Meeting modals
      if (els.miModalEl && header.contains(els.miModalEl)) {
        document.body.appendChild(els.miModalEl);
      }
      if (els.miDetailsModalEl && header.contains(els.miDetailsModalEl)) {
        document.body.appendChild(els.miDetailsModalEl);
      }
      // Notice of Meeting modals
      if (els.noModalEl && header.contains(els.noModalEl)) {
        document.body.appendChild(els.noModalEl);
      }
      if (els.noDetailsModalEl && header.contains(els.noDetailsModalEl)) {
        document.body.appendChild(els.noDetailsModalEl);
      }
    }
  } catch (_) { }

  const STORAGE_KEY = 'docreg_outgoing_letters_v1';
  const MO_KEY = 'docreg_memoranda_v1';
  const OO_KEY = 'docreg_office_orders_v1';
  const TO_KEY = 'docreg_travel_orders_v1';
  const MI_KEY = 'docreg_minutes_v1';
  let olModal, detailsModal;
  try { olModal = new bootstrap.Modal(els.modalEl); } catch (_) { olModal = { show() { }, hide() { } }; }
  try { detailsModal = new bootstrap.Modal(els.detailsModalEl); } catch (_) { detailsModal = { show() { }, hide() { } }; }
  // Incoming
  let ilModal, ilDetailsModal;
  try { ilModal = new bootstrap.Modal(els.ilModalEl); } catch (_) { ilModal = { show() { }, hide() { } }; }
  try { ilDetailsModal = new bootstrap.Modal(els.ilDetailsModalEl); } catch (_) { ilDetailsModal = { show() { }, hide() { } }; }
  // Memorandum
  let moModal, moDetailsModal;
  try { moModal = new bootstrap.Modal(els.moModalEl); } catch (_) { moModal = { show() { }, hide() { } }; }
  try { moDetailsModal = new bootstrap.Modal(els.moDetailsModalEl); } catch (_) { moDetailsModal = { show() { }, hide() { } }; }
  // Office Order
  let ooModal, ooDetailsModal;
  try { ooModal = new bootstrap.Modal(els.ooModalEl); } catch (_) { ooModal = { show() { }, hide() { } }; }
  try { ooDetailsModal = new bootstrap.Modal(els.ooDetailsModalEl); } catch (_) { ooDetailsModal = { show() { }, hide() { } }; }
  // Travel Order
  let toModal, toDetailsModal;
  try { toModal = new bootstrap.Modal(els.toModalEl); } catch (_) { toModal = { show() { }, hide() { } }; }
  try { toDetailsModal = new bootstrap.Modal(els.toDetailsModalEl); } catch (_) { toDetailsModal = { show() { }, hide() { } }; }
  // Minutes of Meeting
  let miModal, miDetailsModal;
  try { miModal = new bootstrap.Modal(els.miModalEl); } catch (_) { miModal = { show() { }, hide() { } }; }
  try { miDetailsModal = new bootstrap.Modal(els.miDetailsModalEl); } catch (_) { miDetailsModal = { show() { }, hide() { } }; }
  // Notice of Meeting
  let noModal, noDetailsModal;
  try { noModal = new bootstrap.Modal(els.noModalEl); } catch (_) { noModal = { show() { }, hide() { } }; }
  try { noDetailsModal = new bootstrap.Modal(els.noDetailsModalEl); } catch (_) { noDetailsModal = { show() { }, hide() { } }; }
  let replyChoiceModal;
  try { replyChoiceModal = new bootstrap.Modal(document.getElementById('replyChoiceModal')); } catch (_) { replyChoiceModal = { show() { }, hide() { } }; }
  let replyContext = null; // { incomingId }

  // ---- Admin detection (only Admin can delete) ----
  const ADMIN_EMAIL = "johnlowel.fradejas@mwss.gov.ph";
  const ADMIN_EMAIL_LOWER = ADMIN_EMAIL.toLowerCase();
  let isAdmin = false;
  // Department-based editor access (non-listed departments are view-only)
  const ALLOWED_DOCREG_DEPTS = ['FOMD', 'PMO', 'SOMD', 'WSMD', 'TIRDD', 'WMD', 'ETOG'];
  let userDepartment = '';
  let canEditDocReg = false;       // governs Add/Edit visibility and actions
  let canAccessDocLinks = false;   // governs clickable links in details
  function computeIsAdmin() {
    try {
      const u = window.firebase?.auth?.().currentUser;
      const em = (u && u.email ? u.email : '').trim().toLowerCase();
      isAdmin = (em === ADMIN_EMAIL_LOWER);
    } catch (_) { isAdmin = false; }
  }
  async function computeDeptAccess() {
    userDepartment = '';
    try {
      const db = window.firebase?.firestore?.();
      const u = window.firebase?.auth?.().currentUser;
      const emailExact = (u?.email || '').trim();
      const emailLower = emailExact.toLowerCase();
      const uid = u?.uid || '';
      if (db) {
        // 1) Prefer UID document (canonical location in main app)
        try {
          if (uid) {
            const byUid = await db.collection('users').doc(uid).get();
            if (byUid.exists) { userDepartment = (byUid.data()?.department || '').trim(); }
          }
        } catch (_) { }
        // 2) Fallback: exact email match
        if (!userDepartment && emailExact) {
          try {
            const snap1 = await db.collection('users').where('email', '==', emailExact).limit(1).get();
            if (!snap1.empty) { userDepartment = (snap1.docs[0].data()?.department || '').trim(); }
          } catch (_) { }
        }
        // 3) Fallback: lowercase email match (in case stored normalized)
        if (!userDepartment && emailLower) {
          try {
            const snap2 = await db.collection('users').where('email', '==', emailLower).limit(1).get();
            if (!snap2.empty) { userDepartment = (snap2.docs[0].data()?.department || '').trim(); }
          } catch (_) { }
        }
      }
    } catch (_) { }
    const deptOk = ALLOWED_DOCREG_DEPTS.map(s => s.toLowerCase()).includes((userDepartment || '').toLowerCase());
    canEditDocReg = !!(isAdmin || deptOk);
    canAccessDocLinks = !!(isAdmin || deptOk);
  }
  function applyRoleUI() {
    // Toggle Add buttons and Reply buttons
    try { if (els.addBtn) els.addBtn.style.display = canEditDocReg ? 'inline-block' : 'none'; } catch (_) { }
    try { if (els.ilAddBtn) els.ilAddBtn.style.display = canEditDocReg ? 'inline-block' : 'none'; } catch (_) { }
    try { if (els.moAddBtn) els.moAddBtn.style.display = canEditDocReg ? 'inline-block' : 'none'; } catch (_) { }
    try { if (els.ooAddBtn) els.ooAddBtn.style.display = canEditDocReg ? 'inline-block' : 'none'; } catch (_) { }
    try { if (els.toAddBtn) els.toAddBtn.style.display = canEditDocReg ? 'inline-block' : 'none'; } catch (_) { }
    try { if (els.miAddBtn) els.miAddBtn.style.display = canEditDocReg ? 'inline-block' : 'none'; } catch (_) { }
    try { if (els.noAddBtn) els.noAddBtn.style.display = canEditDocReg ? 'inline-block' : 'none'; } catch (_) { }
    try { const rb = document.getElementById('ilReplyBtn'); if (rb) rb.style.display = canEditDocReg ? 'inline-block' : 'none'; } catch (_) { }
  }
  // Helper to render doc links respecting access
  function renderDocLink(url) {
    const link = (url || '').trim();
    if (!link) return '';
    if (canAccessDocLinks) {
      return `<a href="${escapeAttr(link)}" target="_blank" rel="noopener">${escapeHtml(link)}</a>`;
    }
    return `<span class="text-muted">Restricted Access</span>`;
  }

  // After details HTML is injected, enforce link restrictions defensively
  function restrictDocLinksIn(container) {
    try {
      if (canAccessDocLinks) return;
      if (!container) return;
      // Replace known Document Link value with Restricted Access
      const labels = Array.from(container.querySelectorAll('.small.text-muted'));
      const docLabel = labels.find(el => (el.textContent || '').trim() === 'Document Link');
      if (docLabel && docLabel.parentElement) {
        const val = docLabel.parentElement.querySelector('.fw-semibold');
        if (val) { val.textContent = 'Restricted Access'; }
      }
      // Disable/replace any anchors inside details
      container.querySelectorAll('a').forEach(a => {
        try { a.addEventListener('click', (e) => e.preventDefault(), { once: true }); } catch (_) { }
        a.removeAttribute('href');
        a.classList.add('text-muted');
        a.textContent = 'Restricted Access';
      });
    } catch (_) { }
  }
  try {
    if (!window.ilUnlinkReply) {
      window.ilUnlinkReply = async function (incomingId, replyId) {
        try {
          const items = ilLoad();
          const idx = items.findIndex(x => x.id === incomingId);
          if (idx >= 0) {
            const r = (items[idx].replies || []).filter(x => x && x.id !== replyId);
            const unlinkedSet = new Set(Array.isArray(items[idx].unlinkedReplyIds) ? items[idx].unlinkedReplyIds : []);
            unlinkedSet.add(replyId);
            const unlinkedArr = Array.from(unlinkedSet);
            items[idx] = { ...items[idx], replies: r, unlinkedReplyIds: unlinkedArr };
            ilSave(items);
            try {
              const db = window.firebase?.firestore?.();
              if (db) {
                const FieldValue = window.firebase?.firestore?.FieldValue;
                const payload = FieldValue ? { replies: r, unlinkedReplyIds: FieldValue.arrayUnion(replyId) } : { replies: r, unlinkedReplyIds: unlinkedArr };
                db.collection('docreg_incoming').doc(incomingId).set(payload, { merge: true });
              }
            } catch (_) { }
          }
        } catch (_) { }
        try {
          const outs = loadItems();
          const oIdx = outs.findIndex(x => x.id === replyId);
          if (oIdx >= 0 && outs[oIdx].replyToIncomingId === incomingId) { const obj = { ...outs[oIdx] }; delete obj.replyToIncomingId; outs[oIdx] = obj; saveItems(outs); }
        } catch (_) { }
        try {
          const mem = moLoad();
          const mIdx = mem.findIndex(x => x.id === replyId);
          if (mIdx >= 0 && mem[mIdx].replyToIncomingId === incomingId) { const obj = { ...mem[mIdx] }; delete obj.replyToIncomingId; mem[mIdx] = obj; moSave(mem); }
        } catch (_) { }
        try {
          const db = window.firebase?.firestore?.();
          if (db) {
            const FieldValue = window.firebase?.firestore?.FieldValue;
            const collections = ['docreg_outgoing', 'docreg_memoranda', 'docreg_notices'];
            for (const col of collections) {
              try {
                const ref = db.collection(col).doc(replyId);
                const snap = await ref.get();
                if (snap.exists) {
                  if (FieldValue) {
                    await ref.set({ replyToIncomingId: FieldValue.delete() }, { merge: true });
                  } else {
                    await ref.set({ replyToIncomingId: null }, { merge: true });
                  }
                }
              } catch (_) { }
            }
          }
        } catch (_) { }
        try { ilOpenDetails(incomingId); } catch (_) { }
      };
    }
  } catch (_) { }
  try {
    if (!window.ilRemoveUnlinked) {
      window.ilRemoveUnlinked = function (incomingId, replyId) {
        if (!incomingId || !replyId) return;
        try {
          const items = ilLoad();
          const idx = items.findIndex(x => x.id === incomingId);
          if (idx >= 0) {
            const arr = Array.isArray(items[idx].unlinkedReplyIds) ? items[idx].unlinkedReplyIds.filter(id => id !== replyId) : [];
            items[idx] = { ...items[idx], unlinkedReplyIds: arr };
            ilSave(items);
          }
        } catch (_) { }
        try {
          const FieldValue = window.firebase?.firestore?.FieldValue;
          const db = window.firebase?.firestore?.();
          if (db && FieldValue) { db.collection('docreg_incoming').doc(incomingId).set({ unlinkedReplyIds: FieldValue.arrayRemove(replyId) }, { merge: true }); }
        } catch (_) { }
      };
    }
  } catch (_) { }
  try {
    window.firebase?.auth?.().onAuthStateChanged(async () => {
      computeIsAdmin();
      await computeDeptAccess();
      applyRoleUI();
      // Refresh all sections so action buttons reflect role
      try { render(); } catch (_) { }
      try { renderIl(); } catch (_) { }
      try { renderMo(); } catch (_) { }
      try { renderOo(); } catch (_) { }
      try { renderTo(); } catch (_) { }
      try { renderMi(); } catch (_) { }
      try { renderNo(); } catch (_) { }
    });
  } catch (_) { computeIsAdmin(); }

  function loadItems() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch (_) { return []; }
  }
  function saveItems(items) {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(items || [])); } catch (_) { }
  }
  function moLoad() { try { const raw = localStorage.getItem(MO_KEY); const arr = raw ? JSON.parse(raw) : []; return Array.isArray(arr) ? arr : []; } catch (_) { return []; } }
  function moSave(items) { try { localStorage.setItem(MO_KEY, JSON.stringify(items || [])); } catch (_) { } }
  function ooLoad() { try { const raw = localStorage.getItem(OO_KEY); const arr = raw ? JSON.parse(raw) : []; return Array.isArray(arr) ? arr : []; } catch (_) { return []; } }
  function ooSave(items) { try { localStorage.setItem(OO_KEY, JSON.stringify(items || [])); } catch (_) { } }
  function toLoad() { try { const raw = localStorage.getItem(TO_KEY); const arr = raw ? JSON.parse(raw) : []; return Array.isArray(arr) ? arr : []; } catch (_) { return []; } }
  function toSave(items) { try { localStorage.setItem(TO_KEY, JSON.stringify(items || [])); } catch (_) { } }
  function uid() { return 'id_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36); }
  // Render date as 'DD Mon YYYY' without timezone shifts. Accepts 'YYYY-MM-DD' or ISO strings.
  function fmtDate(ymd) {
    if (!ymd) return '';
    const s = String(ymd).trim();
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
    if (!m) return s;
    const y = m[1];
    const mm = parseInt(m[2], 10);
    const dd = m[3];
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const mon = months[mm - 1] || m[2];
    return `${dd} ${mon} ${y}`;
  }
  // Format HH:MM or HH:MM:SS to 12-hour time with AM/PM; if already AM/PM, pass-through
  function fmtTime(hhmm) {
    try {
      const s = (hhmm || '').trim();
      if (!s) return '';
      if (/[ap]\.?m\.?$/i.test(s)) return s.toUpperCase().replace(/\./g, '');
      const m = /^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$/.exec(s);
      if (!m) return s;
      let h = parseInt(m[1], 10); const min = m[2];
      const ampm = h >= 12 ? 'PM' : 'AM';
      h = h % 12; if (h === 0) h = 12;
      return `${h}:${min} ${ampm}`;
    } catch (_) { return hhmm || ''; }
  }
  function normalizeYmd(s) {
    if (!s) return '';
    const m = /^\s*(\d{4})-(\d{1,2})-(\d{1,2})\s*$/.exec(s);
    if (!m) return '';
    const y = m[1];
    const mo = String(parseInt(m[2], 10)).padStart(2, '0');
    const da = String(parseInt(m[3], 10)).padStart(2, '0');
    return `${y}-${mo}-${da}`;
  }
  const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn.apply(null, a), ms); } };

  // Working days (Mon-Fri) difference between two YYYY-MM-DD dates, excluding the start day (LOCAL time)
  function businessDaysDiff(startYmd, endYmd) {
    if (!startYmd || !endYmd) return NaN;
    const m1 = /^\s*(\d{4})-(\d{2})-(\d{2})\s*$/.exec(startYmd);
    const m2 = /^\s*(\d{4})-(\d{2})-(\d{2})\s*$/.exec(endYmd);
    if (!m1 || !m2) return NaN;
    const s = new Date(parseInt(m1[1], 10), parseInt(m1[2], 10) - 1, parseInt(m1[3], 10));
    const e = new Date(parseInt(m2[1], 10), parseInt(m2[2], 10) - 1, parseInt(m2[3], 10));
    if (isNaN(s) || isNaN(e)) return NaN;
    if (e < s) return 0;
    let days = 0;
    const cur = new Date(s.getFullYear(), s.getMonth(), s.getDate() + 1); // day after start
    while (cur <= e) { const wd = cur.getDay(); if (wd !== 0 && wd !== 6) days++; cur.setDate(cur.getDate() + 1); }
    return days;
  }

  function resetForm() {
    els.form.reset();
  }

  function rowHtml(it) {
    const name = it.toName || it.name || '';
    const designation = it.toDesignation || it.designation || '';
    const agency = it.toAgency || it.agency || '';
    const address = it.toAddress || it.address || '';
    const addrBlock = `
      <div class="fw-semibold">${escapeHtml(name)}</div>
      <div class="text-muted small">${escapeHtml(designation)}</div>
      <div class="text-muted small">${escapeHtml(agency)}</div>
      <div class="text-muted small">${escapeHtml(address)}</div>
    `;
    const editBtn = canEditDocReg ? `<button class="btn btn-sm btn-outline-primary me-1 ol-edit" data-id="${it.id}" onclick="event.stopPropagation(); window.olEdit && window.olEdit('${it.id}')"><i class="fa fa-pen"></i></button>` : '';
    const delBtn = isAdmin ? `<button class="btn btn-sm btn-outline-danger ol-del" data-id="${it.id}" onclick="event.stopPropagation(); window.olDelete && window.olDelete('${it.id}')"><i class="fa fa-trash"></i></button>` : '';
    return `<tr data-id="${it.id}" class="ol-row" onclick="window.olOpenDetails && window.olOpenDetails(this.getAttribute('data-id'))">
      <td>${escapeHtml((it.controlNo || it.refNo) || '')}</td>
      <td>${fmtDate(it.docDate) || ''}</td>
      <td>${addrBlock}</td>
      <td>${escapeHtml(it.subject || '')}</td>
      <td>${escapeHtml(it.preparedBy || '')}</td>
      <td class="text-nowrap">${editBtn}${delBtn}</td>
    </tr>`;
  }

  const PAGE_SIZE = 60;

  let currentPage = 1;

  function buildPagination(total) {
    const ul = document.getElementById('olPagination'); if (!ul) return;
    const maxPage = Math.max(1, Math.ceil(total / PAGE_SIZE));
    if (currentPage > maxPage) currentPage = maxPage;
    let html = '';
    html += `<li class="page-item${currentPage <= 1 ? ' disabled' : ''}"><a class="page-link" href="#" data-page="prev">Prev</a></li>`;
    for (let p = 1; p <= maxPage; p++) {
      html += `<li class="page-item${p === currentPage ? ' active' : ''}"><a class="page-link" href="#" data-page="${p}">${p}</a></li>`;
    }
    html += `<li class="page-item${currentPage >= maxPage ? ' disabled' : ''}"><a class="page-link" href="#" data-page="next">Next</a></li>`;
    ul.innerHTML = html;
    ul.onclick = (ev) => {
      const a = ev.target.closest('a[data-page]'); if (!a) return;
      ev.preventDefault();
      const val = a.getAttribute('data-page');
      if (val === 'prev' && currentPage > 1) currentPage--;
      else if (val === 'next') { const max = Math.max(1, Math.ceil(total / PAGE_SIZE)); if (currentPage < max) currentPage++; }
      else currentPage = parseInt(val, 10) || 1;
      render();
    };
  }

  function render() {
    const items = loadItems().slice();
    // Rebuild delivery filter options from items (unique, non-empty)
    try {
      const uniq = Array.from(new Set(items.map(i => (i.deliveryVia || '').trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b));
      if (els.deliveryFilter) {
        const curr = els.deliveryFilter.value || '';
        const html = ['<option value="">All Methods</option>'].concat(uniq.map(v => `<option${v === curr ? ' selected' : ''}>${escapeHtml(v)}</option>`)).join('');
        els.deliveryFilter.innerHTML = html;
        // If previous value no longer exists, keep as All Methods
      }
    } catch (_) { }

    // Apply filters
    const text = (els.searchInput?.value || '').toLowerCase().trim();
    const status = els.statusFilter?.value || '';
    const via = els.deliveryFilter?.value || '';
    let list = items;
    if (status) list = list.filter(it => (it.status || '') === status);
    if (via) list = list.filter(it => (it.deliveryVia || '') === via);
    if (text) {
      list = list.filter(it => {
        const hay = [
          it.controlNo, it.refNo,
          it.toName, it.toDesignation, it.toAgency, it.toAddress,
          it.name, it.designation, it.agency, it.address,
          it.subject, it.preparedBy, it.signatory
        ].map(v => (v || '').toLowerCase()).join(' \u0000 ');
        return hay.includes(text);
      });
    }
    // Sort by Document Date desc then by createdAt desc (latest on top)
    list.sort((a, b) => {
      const ad = a.docDate || ''; const bd = b.docDate || '';
      if (ad !== bd) return ad < bd ? 1 : -1;
      const ac = a.createdAt || ''; const bc = b.createdAt || '';
      return ac < bc ? 1 : -1;
    });
    const total = list.length;
    const start = (currentPage - 1) * PAGE_SIZE;
    const end = Math.min(start + PAGE_SIZE, total);
    const slice = list.slice(start, end);
    els.tbody.innerHTML = slice.map(rowHtml).join('');
    els.empty.style.display = total ? 'none' : 'block';
    buildPagination(total);
  }

  function onAddClick() {
    if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    resetForm();
    // default dates
    try { els.docDate.value = new Date().toISOString().slice(0, 10); } catch (_) { }
    try { if (els.signatory) els.signatory.value = ''; } catch (_) { }
    olModal.show();
  }
  function prefillMemoFromIncoming(incoming) {
    // Switch to Memorandum section and open the Memorandum modal prefilled
    switchSection('memo');
    moResetForm();
    if (els.moRefNo) els.moRefNo.value = `${incoming.refNo || ''}`;
    if (els.moDocDate) els.moDocDate.value = new Date().toISOString().slice(0, 10);
    // New structured TO/FROM
    if (els.moToName) els.moToName.value = incoming.name || '';
    if (els.moToDesignation) els.moToDesignation.value = incoming.designation || '';
    if (els.moToDepartment) els.moToDepartment.value = incoming.agency || '';
    if (els.moFromName) els.moFromName.value = '';
    if (els.moFromDesignation) els.moFromDesignation.value = '';
    if (els.moFromDepartment) els.moFromDepartment.value = '';
    // Back-compat fallback if the old textarea exists in DOM in some builds
    if (els.moPartyBlock) {
      const lines = [incoming.name || '', incoming.designation || '', incoming.agency || '', incoming.address || ''];
      els.moPartyBlock.value = lines.join('\n').replace(/\n+$/, '');
    }
    if (els.moSubject) els.moSubject.value = `RE: ${incoming.subject || ''}`.trim();
    if (els.moPreparedBy) els.moPreparedBy.value = '';
    // Linkage
    els.moForm?.setAttribute('data-reply-to', incoming.id);
    moModal.show();
  }

  function moFinalizeReplyLinkage(createdMemo) {
    const incomingId = els.moForm?.getAttribute('data-reply-to');
    if (!incomingId) return;
    els.moForm.removeAttribute('data-reply-to');
    try {
      // Ensure memo carries replyToIncomingId
      const list = moLoad();
      const idx = list.findIndex(x => x.id === createdMemo.id);
      if (idx >= 0) { list[idx] = { ...list[idx], replyToIncomingId: incomingId }; moSave(list); }
      // Update incoming record replies
      const inItems = ilLoad();
      const i = inItems.findIndex(x => x.id === incomingId);
      if (i >= 0) { const r = inItems[i].replies || []; r.push({ id: createdMemo.id, type: 'memo', date: createdMemo.docDate }); inItems[i] = { ...inItems[i], replies: r }; ilSave(inItems); }
    } catch (_) { }
  }

  function onSubmit(e) {
    e.preventDefault();
    if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    if (!els.form.checkValidity()) {
      els.form.reportValidity();
      return;
    }
    const replyTo = els.form.getAttribute('data-reply-to') || '';
    const partyLines = (els.olPartyBlock?.value || '').split(/\r?\n/).map(s => s.trim());
    const partyName = partyLines[0] || '';
    const partyDesignation = partyLines[1] || '';
    const partyAgency = partyLines[2] || '';
    const partyAddress = partyLines[3] || '';
    // Structured TO/FROM (prefer these; fallback to legacy textarea)
    const toName = (els.olToName?.value || '').trim() || partyName;
    const toPosition = (els.olToPosition?.value || '').trim() || partyDesignation;
    const toAgency = (els.olToAgency?.value || '').trim() || partyAgency;
    const toAddress = (els.olToAddress?.value || '').trim() || partyAddress;
    const fromName = (els.olFromName?.value || '').trim();
    const fromDesignation = (els.olFromDesignation?.value || '').trim();
    const fromAgency = (els.olFromAgency?.value || '').trim();
    const fromAddress = (els.olFromAddress?.value || '').trim();
    const it = {
      id: els.id?.value || uid(),
      refNo: els.refNo.value.trim(),
      controlNo: (els.controlNo?.value || '').trim(),
      docDate: els.docDate.value,
      // Legacy block mirrors TO for display-compat
      name: toName,
      designation: toPosition,
      agency: toAgency,
      address: toAddress,
      // Structured fields
      toName, toPosition, toAgency, toAddress,
      fromName, fromDesignation, fromAgency, fromAddress,
      subject: els.subject.value.trim(),
      preparedBy: els.preparedBy.value.trim(),
      docLink: (els.docLink?.value || '').trim(),
      deliveryVia: els.deliveryVia.value.trim(),
      status: els.status.value,
      remarks: els.remarks.value || '',
      signatory: (els.signatory?.value || '').trim(),
    };
    // Add optional fields only if they have values (avoid undefined in Firestore)
    if (!els.id?.value) { it.createdAt = new Date().toISOString(); }
    if (replyTo) { it.replyToIncomingId = replyTo; }

    const items = loadItems();
    if (els.id?.value) {
      // update existing
      const idx = items.findIndex(x => x.id === it.id);
      if (idx >= 0) { items[idx] = { ...items[idx], ...it }; }
    } else {
      items.push(it);
    }
    saveItems(items);
    try { finalizeReplyLinkage(it); } catch (_) { }
    olModal.hide();
    render();
  }

  function openEdit(id) {
    if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    const items = loadItems();
    const it = items.find(x => x.id === id); if (!it) return;
    resetForm();
    if (els.id) els.id.value = it.id;
    els.refNo.value = it.refNo || '';
    if (els.controlNo) els.controlNo.value = it.controlNo || '';
    els.docDate.value = it.docDate || '';
    if (els.olPartyBlock) {
      const lines = [it.name || '', it.designation || '', it.agency || '', it.address || ''];
      els.olPartyBlock.value = lines.join('\n').replace(/\n+$/, '');
    }
    // Structured TO/FROM prefill
    if (els.olToName) els.olToName.value = (it.toName || it.name || '');
    if (els.olToPosition) els.olToPosition.value = (it.toPosition || it.toDesignation || it.designation || '');
    if (els.olToAgency) els.olToAgency.value = (it.toAgency || it.agency || '');
    if (els.olToAddress) els.olToAddress.value = (it.toAddress || it.address || '');
    if (els.olFromName) els.olFromName.value = (it.fromName || '');
    if (els.olFromDesignation) els.olFromDesignation.value = (it.fromDesignation || '');
    if (els.olFromAgency) els.olFromAgency.value = (it.fromAgency || '');
    if (els.olFromAddress) els.olFromAddress.value = (it.fromAddress || '');
    els.subject.value = it.subject || '';
    if (els.docLink) els.docLink.value = it.docLink || '';
    els.preparedBy.value = it.preparedBy || '';
    els.deliveryVia.value = it.deliveryVia || '';
    els.status.value = it.status || '';
    els.remarks.value = it.remarks || '';
    if (els.signatory) els.signatory.value = it.signatory || '';
    olModal.show();
  }

  function deleteItem(id, opts = {}) {
    const { skipConfirm = false } = opts || {};
    if (!isAdmin) { alert('Only the Administrator can delete.'); return; }
    const items = loadItems();
    const it = items.find(x => x.id === id);
    if (!it) return;
    const ok = skipConfirm ? true : confirm(`Delete document ${it.refNo || it.controlNo || ''}? This cannot be undone.`);
    if (!ok) return;
    const next = items.filter(x => x.id !== id);
    saveItems(next);
    render();
  }

  // Global wrappers for inline action buttons
  window.olEdit = (id) => { try { openEdit(id); } catch (_) { } };
  window.olDelete = (id) => {
    try {
      const items = loadItems();
      const it = items.find(x => x.id === id);
      const label = it ? (it.refNo || it.controlNo || '') : '';
      const ok = confirm(`Delete document${label ? ' ' + label : ''}? This cannot be undone.`);
      if (!ok) return;
      deleteItem(id, { skipConfirm: true });
    } catch (_) { }
  };

  // Expose getter for Outgoing items by id (used by print action in docreg.html)
  window.olGetItem = (id) => {
    try {
      const items = loadItems();
      return items.find(x => x.id === id) || null;
    } catch (_) { return null; }
  };

  function openDetails(id) {
    const items = loadItems();
    const it = items.find(x => x.id === id); if (!it) return;
    try { window.__olCurrentItem = it; } catch (_) { }
    const toName = (it.name || '').trim();
    const toDesignation = (it.designation || '').trim();
    const toAgency = (it.agency || '').trim();
    const toAddress = (it.address || '').trim();
    const letterToParts = [];
    if (toName) letterToParts.push(`<div class="fw-semibold">${escapeHtml(toName)}</div>`);
    if (toDesignation) letterToParts.push(`<div>${escapeHtml(toDesignation)}</div>`);
    if (toAgency) letterToParts.push(`<div>${escapeHtml(toAgency)}</div>`);
    if (toAddress) letterToParts.push(`<div>${escapeHtml(toAddress)}</div>`);
    const letterToHtml = letterToParts.join('');
    const html = `
      <div class="row g-2">
        <div class="col-md-4"><div class="small text-muted">Reference No.</div><div class="fw-semibold">${escapeHtml(it.refNo || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Document Date</div><div class="fw-semibold">${fmtDate(it.docDate) || ''}</div></div>
        <div class="col-md-4"><div class="small text-muted">Prepared by</div><div class="fw-semibold">${escapeHtml(it.preparedBy || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Control No.</div><div class="fw-semibold">${escapeHtml(it.controlNo || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Letter to:</div><div>${letterToHtml}</div></div>
        <div class="col-12"><div class="small text-muted">Subject</div><div class="fw-semibold">${escapeHtml(it.subject || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Document Link</div><div class="fw-semibold">${renderDocLink(it.docLink)}</div></div>
        <div class="col-md-6"><div class="small text-muted">Delivery/Sent via</div><div class="fw-semibold">${escapeHtml(it.deliveryVia || '')}</div></div>
        <div class="col-md-6"><div class="small text-muted">Signatory</div><div class="fw-semibold">${escapeHtml(it.signatory || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Status</div><div class="fw-semibold">${escapeHtml(it.status || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Remarks</div><div class="fw-semibold">${escapeHtml(it.remarks || '')}</div></div>
      </div>
    `;
    if (els.detailsBody) {
      els.detailsBody.innerHTML = html;
      restrictDocLinksIn(els.detailsBody);
      try {
        const __hist = (it.history || []).slice().sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        if (__hist.length) {
          const rows = __hist.map(h => {
            const who = (h.fullName || '').trim() || (h.email || 'unknown');
            const when = (new Date(h.timestamp)).toLocaleString();
            const act = h.action || 'update';
            return `<div class="d-flex align-items-start gap-2 py-1"><i class="fa-regular fa-clock mt-1 text-muted"></i><div><div><strong>${escapeHtml(act)}</strong> by ${escapeHtml(who)}</div><div class="text-muted small">${escapeHtml(when)}</div></div></div>`;
          }).join('');
          els.detailsBody.insertAdjacentHTML('beforeend', `<div class="small text-muted mt-3">Edit History</div><div class="border rounded p-2 bg-light small">${rows}</div>`);
        }
      } catch (_) { }
      try { els.detailsBody.setAttribute('data-current-id', it.id); } catch (_) { }
    }
    try { window.__olCurrentItem = it; } catch (_) { }
    detailsModal.show();
    try { window.ensureOlDetailsPrintButton && window.ensureOlDetailsPrintButton(); } catch (_) { }
    try {
      const modal = els.detailsModalEl || document.getElementById('olDetailsModal');
      const footer = modal ? modal.querySelector('.modal-footer') : null;

      // First check if there's an existing button with id="olPrintBtn" that needs a handler
      const existingPrintBtn = document.getElementById('olPrintBtn');
      if (existingPrintBtn && !existingPrintBtn._hasWordHandler) {
        existingPrintBtn._hasWordHandler = true;
        existingPrintBtn.addEventListener('click', () => {
          try {
            const cur = window.__olCurrentItem || it;
            console.log('Download Word clicked (existing btn), item:', cur);
            if (window.olGenerateFromTemplate) {
              console.log('Calling olGenerateFromTemplate...');
              window.olGenerateFromTemplate(cur);
            } else {
              console.warn('olGenerateFromTemplate not found, opening template');
              window.open('assets/outgoing_letter_template.docx', '_blank');
            }
          } catch (err) {
            console.error('Download Word error:', err);
            alert('Error generating document: ' + (err.message || err));
          }
        });
      }

      // Fallback: if no existing button, create one
      if (footer && !footer.querySelector('.ol-print-btn') && !existingPrintBtn) {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'btn btn-outline-primary ol-print-btn';
        b.innerHTML = '<i class="fa fa-download me-1"></i>Download Word';
        b.addEventListener('click', () => {
          try {
            const cur = window.__olCurrentItem || it;
            console.log('Download Word clicked, item:', cur);
            if (window.olGenerateFromTemplate) {
              console.log('Calling olGenerateFromTemplate...');
              window.olGenerateFromTemplate(cur);
            } else {
              console.warn('olGenerateFromTemplate not found, opening template');
              window.open('assets/outgoing_letter_template.docx', '_blank');
            }
          } catch (err) {
            console.error('Download Word error:', err);
            alert('Error generating document: ' + (err.message || err));
          }
        });
        const closeBtn = footer.querySelector('[data-bs-dismiss="modal"]');
        if (closeBtn) { footer.insertBefore(b, closeBtn); } else { footer.appendChild(b); }
      }
    } catch (_) { }
  }

  // Fallback global hook used by inline onclick on <tr>
  window.olOpenDetails = (id) => { try { openDetails(id); } catch (_) { } };

  // Simple HTML escaper for table content
  function escapeHtml(s) {
    return (s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\'': '&#39;' }[c]));
  }
  function escapeAttr(s) {
    // Minimal attribute escaping and whitespace encoding for URLs
    return (s || '').replace(/["'`<>\s]/g, (c) => {
      const map = { '"': '&quot;', "'": '&#39;', '`': '&#96;', '<': '&lt;', '>': '&gt;', ' ': '%20', '\n': '%0A', '\r': '%0D', '\t': '%09' };
      return map[c] || '';
    });
  }

  // ---- Section switcher ----
  function switchSection(which) {
    // Fade out
    const mainContent = document.getElementById('docMain');
    if (mainContent) mainContent.classList.add('fade-out');

    setTimeout(() => {
      if (which === 'incoming') {
        if (els.outgoingSection) els.outgoingSection.style.display = 'none';
        if (els.memoSection) els.memoSection.style.display = 'none';
        if (els.officeSection) els.officeSection.style.display = 'none';
        if (els.travelSection) els.travelSection.style.display = 'none';
        if (els.minutesSection) els.minutesSection.style.display = 'none';
        if (els.noticeSection) els.noticeSection.style.display = 'none';
        if (els.incomingSection) {
          els.incomingSection.style.display = 'block';
          els.incomingSection.classList.add('fade-element');
        }
        els.btnOutgoing?.classList.remove('active');
        els.btnMemo?.classList.remove('active');
        els.btnOffice?.classList.remove('active');
        els.btnTravel?.classList.remove('active');
        els.btnMinutes?.classList.remove('active');
        els.btnNotice?.classList.remove('active');
        els.btnIncoming?.classList.add('active');
        try { renderIl(); } catch (_) { }
      } else if (which === 'office') {
        if (els.incomingSection) els.incomingSection.style.display = 'none';
        if (els.outgoingSection) els.outgoingSection.style.display = 'none';
        if (els.memoSection) els.memoSection.style.display = 'none';
        if (els.travelSection) els.travelSection.style.display = 'none';
        if (els.minutesSection) els.minutesSection.style.display = 'none';
        if (els.noticeSection) els.noticeSection.style.display = 'none';
        if (els.officeSection) {
          els.officeSection.style.display = 'block';
          els.officeSection.classList.add('fade-element');
        }
        els.btnIncoming?.classList.remove('active');
        els.btnOutgoing?.classList.remove('active');
        els.btnMemo?.classList.remove('active');
        els.btnTravel?.classList.remove('active');
        els.btnMinutes?.classList.remove('active');
        els.btnNotice?.classList.remove('active');
        els.btnOffice?.classList.add('active');
        try { renderOo(); } catch (_) { }
      } else if (which === 'travel') {
        if (els.incomingSection) els.incomingSection.style.display = 'none';
        if (els.outgoingSection) els.outgoingSection.style.display = 'none';
        if (els.memoSection) els.memoSection.style.display = 'none';
        if (els.officeSection) els.officeSection.style.display = 'none';
        if (els.minutesSection) els.minutesSection.style.display = 'none';
        if (els.noticeSection) els.noticeSection.style.display = 'none';
        if (els.travelSection) {
          els.travelSection.style.display = 'block';
          els.travelSection.classList.add('fade-element');
        }
        els.btnIncoming?.classList.remove('active');
        els.btnOutgoing?.classList.remove('active');
        els.btnMemo?.classList.remove('active');
        els.btnOffice?.classList.remove('active');
        els.btnMinutes?.classList.remove('active');
        els.btnNotice?.classList.remove('active');
        els.btnTravel?.classList.add('active');
        try { renderTo(); } catch (_) { }
      } else if (which === 'minutes') {
        if (els.incomingSection) els.incomingSection.style.display = 'none';
        if (els.outgoingSection) els.outgoingSection.style.display = 'none';
        if (els.memoSection) els.memoSection.style.display = 'none';
        if (els.officeSection) els.officeSection.style.display = 'none';
        if (els.travelSection) els.travelSection.style.display = 'none';
        if (els.noticeSection) els.noticeSection.style.display = 'none';
        if (els.minutesSection) {
          els.minutesSection.style.display = 'block';
          els.minutesSection.classList.add('fade-element');
        }
        els.btnIncoming?.classList.remove('active');
        els.btnOutgoing?.classList.remove('active');
        els.btnMemo?.classList.remove('active');
        els.btnOffice?.classList.remove('active');
        els.btnTravel?.classList.remove('active');
        els.btnNotice?.classList.remove('active');
        els.btnMinutes?.classList.add('active');
        try { renderMi(); } catch (_) { }
      } else if (which === 'notice') {
        if (els.incomingSection) els.incomingSection.style.display = 'none';
        if (els.outgoingSection) els.outgoingSection.style.display = 'none';
        if (els.memoSection) els.memoSection.style.display = 'none';
        if (els.officeSection) els.officeSection.style.display = 'none';
        if (els.travelSection) els.travelSection.style.display = 'none';
        if (els.minutesSection) els.minutesSection.style.display = 'none';
        if (els.noticeSection) {
          els.noticeSection.style.display = 'block';
          els.noticeSection.classList.add('fade-element');
        }
        els.btnIncoming?.classList.remove('active');
        els.btnOutgoing?.classList.remove('active');
        els.btnMemo?.classList.remove('active');
        els.btnOffice?.classList.remove('active');
        els.btnTravel?.classList.remove('active');
        els.btnMinutes?.classList.remove('active');
        els.btnNotice?.classList.add('active');
        try { renderNo(); } catch (_) { }
      } else if (which === 'memo') {
        if (els.incomingSection) els.incomingSection.style.display = 'none';
        if (els.outgoingSection) els.outgoingSection.style.display = 'none';
        if (els.officeSection) els.officeSection.style.display = 'none';
        if (els.travelSection) els.travelSection.style.display = 'none';
        if (els.minutesSection) els.minutesSection.style.display = 'none';
        if (els.noticeSection) els.noticeSection.style.display = 'none';
        if (els.memoSection) {
          els.memoSection.style.display = 'block';
          els.memoSection.classList.add('fade-element');
        }
        els.btnIncoming?.classList.remove('active');
        els.btnOutgoing?.classList.remove('active');
        els.btnOffice?.classList.remove('active');
        els.btnTravel?.classList.remove('active');
        els.btnMinutes?.classList.remove('active');
        els.btnNotice?.classList.remove('active');
        els.btnMemo?.classList.add('active');
        try { renderMo(); } catch (_) { }
      } else {
        if (els.incomingSection) els.incomingSection.style.display = 'none';
        if (els.memoSection) els.memoSection.style.display = 'none';
        if (els.officeSection) els.officeSection.style.display = 'none';
        if (els.travelSection) els.travelSection.style.display = 'none';
        if (els.minutesSection) els.minutesSection.style.display = 'none';
        if (els.noticeSection) els.noticeSection.style.display = 'none';
        if (els.outgoingSection) {
          els.outgoingSection.style.display = 'block';
          els.outgoingSection.classList.add('fade-element');
        }
        els.btnIncoming?.classList.remove('active');
        els.btnMemo?.classList.remove('active');
        els.btnOffice?.classList.remove('active');
        els.btnTravel?.classList.remove('active');
        els.btnMinutes?.classList.remove('active');
        els.btnNotice?.classList.remove('active');
        els.btnOutgoing?.classList.add('active');
        try { render(); } catch (_) { }
      }
      // Fade in
      if (mainContent) mainContent.classList.remove('fade-out');
    }, 200);
  }
  if (els.btnOutgoing) els.btnOutgoing.addEventListener('click', (e) => { e.preventDefault(); switchSection('outgoing'); });
  if (els.btnIncoming) els.btnIncoming.addEventListener('click', (e) => { e.preventDefault(); switchSection('incoming'); });
  if (els.btnMemo) els.btnMemo.addEventListener('click', (e) => { e.preventDefault(); switchSection('memo'); });
  if (els.btnOffice) els.btnOffice.addEventListener('click', (e) => { e.preventDefault(); switchSection('office'); });
  if (els.btnTravel) els.btnTravel.addEventListener('click', (e) => { e.preventDefault(); switchSection('travel'); });
  if (els.btnMinutes) els.btnMinutes.addEventListener('click', (e) => { e.preventDefault(); switchSection('minutes'); });
  if (els.btnNotice) els.btnNotice.addEventListener('click', (e) => { e.preventDefault(); switchSection('notice'); });

  // ---- Notice of Meeting module ----
  const NO_PAGE_SIZE = 60;
  let noCurrentPage = 1;

  function noRowHtml(it) {
    const nameAgency = `
      <div class="fw-semibold">${escapeHtml(it.name || '')}</div>
      <div class="text-muted small">${escapeHtml(it.agency || '')}</div>
    `;
    const delBtn = isAdmin ? `<button class=\"btn btn-sm btn-outline-danger no-del\" data-id=\"${it.id}\" onclick=\"event.stopPropagation(); window.noDelete && window.noDelete('${it.id}')\"><i class=\"fa fa-trash\"></i></button>` : '';
    const editBtn = canEditDocReg ? `<button class=\"btn btn-sm btn-outline-primary me-1 no-edit\" data-id=\"${it.id}\" onclick=\"event.stopPropagation(); window.noEdit && window.noEdit('${it.id}')\"><i class=\"fa fa-pen\"></i></button>` : '';
    return `<tr data-id="${it.id}" class="no-row" onclick="window.noOpenDetails && window.noOpenDetails(this.getAttribute('data-id'))">
      <td>${escapeHtml(it.controlNo || '')}</td>
      <td>${fmtDate(it.docDate) || ''}</td>
      <td>${nameAgency}</td>
      <td>${escapeHtml(it.subject || '')}</td>
      <td>${escapeHtml(it.signatory || '')}</td>
      <td class="text-nowrap">${editBtn}${delBtn}</td>
    </tr>`;
  }

  function noBuildPagination(total) {
    const ul = document.getElementById('noPagination'); if (!ul) return;
    const maxPage = Math.max(1, Math.ceil(total / NO_PAGE_SIZE));
    if (noCurrentPage > maxPage) noCurrentPage = maxPage;
    let html = '';
    html += `<li class="page-item${noCurrentPage <= 1 ? ' disabled' : ''}"><a class="page-link" href="#" data-page="prev">Prev</a></li>`;
    for (let p = 1; p <= maxPage; p++) html += `<li class="page-item${p === noCurrentPage ? ' active' : ''}"><a class="page-link" href="#" data-page="${p}">${p}</a></li>`;
    html += `<li class="page-item${noCurrentPage >= maxPage ? ' disabled' : ''}"><a class="page-link" href="#" data-page="next">Next</a></li>`;
    ul.innerHTML = html;
    ul.onclick = (ev) => {
      const a = ev.target.closest('a[data-page]'); if (!a) return; ev.preventDefault();
      const val = a.getAttribute('data-page');
      if (val === 'prev' && noCurrentPage > 1) noCurrentPage--;
      else if (val === 'next') { const max = Math.max(1, Math.ceil(total / NO_PAGE_SIZE)); if (noCurrentPage < max) noCurrentPage++; }
      else noCurrentPage = parseInt(val, 10) || 1;
      renderNo();
    };
  }

  function noLoad() { return []; }

  function renderNo() {
    const items = noLoad().slice();
    const text = (els.noSearchInput?.value || '').toLowerCase().trim();
    let list = items;
    if (text) {
      list = list.filter(it => {
        const hay = [it.controlNo, it.name, it.agency, it.subject, it.signatory, it.venue]
          .map(v => (v || '').toLowerCase()).join(' \u0000 ');
        return hay.includes(text);
      });
    }
    list.sort((a, b) => { const ad = a.docDate || '', bd = b.docDate || ''; if (ad !== bd) return ad < bd ? 1 : -1; const ac = a.createdAt || '', bc = b.createdAt || ''; return ac < bc ? 1 : -1; });
    const total = list.length;
    const start = (noCurrentPage - 1) * NO_PAGE_SIZE;
    const end = Math.min(start + NO_PAGE_SIZE, total);
    const slice = list.slice(start, end);
    if (els.noTbody) els.noTbody.innerHTML = slice.map(noRowHtml).join('');
    if (els.noEmpty) els.noEmpty.style.display = total ? 'none' : 'block';
    noBuildPagination(total);
  }

  function noResetForm() { els.noForm?.reset(); try { if (els.noDocDate) els.noDocDate.value = new Date().toISOString().slice(0, 10); } catch (_) { } }
  function noOnAdd() { noResetForm(); if (els.noId) els.noId.value = ''; noModal.show(); }
  function noOnAdd() { if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; } noResetForm(); if (els.noId) els.noId.value = ''; noModal.show(); }
  function noOnSubmit(e) {
    e.preventDefault(); if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; } if (!els.noForm.checkValidity()) { els.noForm.reportValidity(); return; }
    const lines = (els.noPartyBlock?.value || '').split(/\r?\n/).map(s => s.trim());
    const sigBlock = (els.noSignatoryBlock?.value || '').split(/\r?\n/).map(s => s.trim());
    const it = {
      id: els.noId?.value || ('id_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36)),
      refNo: (els.noRefNo?.value || '').trim(),
      controlNo: (els.noControlNo?.value || '').trim(),
      docDate: els.noDocDate?.value || '',
      meetDate: els.noMeetDate?.value || '',
      meetTime: els.noMeetTime?.value || '',
      name: lines[0] || '',
      agency: lines[1] || '',
      subject: (els.noSubject?.value || '').trim(),
      signatory: (els.noSignatory?.value || '').trim(),
      signatoryDesignation: (sigBlock[0] || (els.noSignatoryDesignation?.value || '').trim()),
      signatoryDepartment: (sigBlock[1] || ''),
      venue: (els.noVenue?.value || '').trim(),
      docLink: (els.noDocLink?.value || '').trim(),
      remarks: (els.noRemarks?.value || '').trim(),
      createdAt: els.noId?.value ? undefined : new Date().toISOString(),
    };
    // By default, rely on Firestore rebinding below; if not available, keep local no-op behavior
    try { window.alert('Offline save is not supported for Notice in this build.'); } catch (_) { }
  }
  function noOpenEdit(id) { const it = noLoad().find(x => x.id === id); if (!it) return; noResetForm(); if (els.noId) els.noId.value = it.id; if (els.noControlNo) els.noControlNo.value = it.controlNo || ''; if (els.noDocDate) els.noDocDate.value = it.docDate || ''; if (els.noPartyBlock) { const lines = [it.name || '', it.agency || '']; els.noPartyBlock.value = lines.join('\n').replace(/\n+$/, ''); } if (els.noSubject) els.noSubject.value = it.subject || ''; if (els.noSignatory) els.noSignatory.value = it.signatory || ''; if (els.noSignatoryDesignation) els.noSignatoryDesignation.value = it.signatoryDesignation || it.designation || ''; if (els.noSignatoryBlock) { const s1 = (it.signatoryDesignation || it.designation || ''); const s2 = (it.signatoryDepartment || ''); els.noSignatoryBlock.value = [s1, s2].join('\n').replace(/\n+$/, ''); } if (els.noVenue) els.noVenue.value = it.venue || ''; if (els.noDocLink) els.noDocLink.value = it.docLink || ''; if (els.noRemarks) els.noRemarks.value = it.remarks || ''; noModal.show(); }
  function noOpenEdit(id) { const it = noLoad().find(x => x.id === id); if (!it) return; noResetForm(); if (els.noId) els.noId.value = it.id; if (els.noRefNo) els.noRefNo.value = it.refNo || ''; if (els.noControlNo) els.noControlNo.value = it.controlNo || ''; if (els.noDocDate) els.noDocDate.value = it.docDate || ''; if (els.noMeetDate) els.noMeetDate.value = it.meetDate || ''; if (els.noMeetTime) els.noMeetTime.value = it.meetTime || ''; if (els.noPartyBlock) { const lines = [it.name || '', it.agency || '']; els.noPartyBlock.value = lines.join('\n').replace(/\n+$/, ''); } if (els.noSubject) els.noSubject.value = it.subject || ''; if (els.noSignatory) els.noSignatory.value = it.signatory || ''; if (els.noSignatoryDesignation) els.noSignatoryDesignation.value = it.signatoryDesignation || it.designation || ''; if (els.noSignatoryBlock) { const s1 = (it.signatoryDesignation || it.designation || ''); const s2 = (it.signatoryDepartment || ''); els.noSignatoryBlock.value = [s1, s2].join('\n').replace(/\n+$/, ''); } if (els.noVenue) els.noVenue.value = it.venue || ''; if (els.noDocLink) els.noDocLink.value = it.docLink || ''; if (els.noRemarks) els.noRemarks.value = it.remarks || ''; noModal.show(); }
  function noDelete(id) { /* re-bound to Firestore below */ }
  function noOpenDetails(id) {
    const it = noLoad().find(x => x.id === id); if (!it) return;
    const html = `
      <div class="row g-2">
        <div class="col-md-4"><div class="small text-muted">Control No.</div><div class="fw-semibold">${escapeHtml(it.controlNo || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Document Date</div><div class="fw-semibold">${fmtDate(it.docDate) || ''}</div></div>
        <div class="col-md-4"><div class="small text-muted">Signatory</div><div class="fw-semibold">${escapeHtml(it.signatory || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Designation</div><div class="fw-semibold">${escapeHtml(it.signatoryDesignation || it.designation || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Department</div><div class="fw-semibold">${escapeHtml(it.signatoryDepartment || '')}</div></div>
        ${it.meetDate ? `<div class="col-md-4" id="noMeetDateRow"><div class="small text-muted">Date of Meeting</div><div class="fw-semibold">${fmtDate(it.meetDate) || ''}</div></div>` : ''}
        ${it.meetTime ? `<div class="col-md-4" id="noMeetTimeRow"><div class="small text-muted">Time of Meeting</div><div class="fw-semibold">${escapeHtml(fmtTime(it.meetTime))}</div></div>` : ''}
        <div class="col-md-6"><div class="small text-muted">Name</div><div class="fw-semibold">${escapeHtml(it.name || '')}</div></div>
        <div class="col-md-6"><div class="small text-muted">Agency</div><div class="fw-semibold">${escapeHtml(it.agency || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Subject</div><div class="fw-semibold">${escapeHtml(it.subject || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Document Link</div><div class="fw-semibold">${it.docLink ? `<a href="${escapeAttr(it.docLink)}" target="_blank" rel="noopener">${escapeHtml(it.docLink)}</a>` : ''}</div></div>
        <div class="col-12"><div class="small text-muted">Remarks</div><div class="fw-semibold">${escapeHtml(it.remarks || '')}</div></div>
      </div>
    `;
    if (els.noDetailsBody) {
      els.noDetailsBody.innerHTML = html;
      try { els.noDetailsBody.setAttribute('data-current-id', it.id); } catch (_) { }
      try { els.noDetailsBody.setAttribute('data-current-json', encodeURIComponent(JSON.stringify(it))); } catch (_) { }
      try { window.__noCurrentItem = it; } catch (_) { }
      try { window.__noCurrentId = it.id; } catch (_) { }
      try {
        const __hist = (it.history || []).slice().sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        if (__hist.length) {
          const rows = __hist.map(h => {
            const who = (h.fullName || '').trim() || (h.email || 'unknown');
            const when = (new Date(h.timestamp)).toLocaleString();
            const act = h.action || 'update';
            return `<div class="d-flex align-items-start gap-2 py-1"><i class="fa-regular fa-clock mt-1 text-muted"></i><div><div><strong>${escapeHtml(act)}</strong> by ${escapeHtml(who)}</div><div class="text-muted small">${escapeHtml(when)}</div></div></div>`;
          }).join('');
          els.noDetailsBody.insertAdjacentHTML('beforeend', `<div class="small text-muted mt-3">Edit History</div><div class="border rounded p-2 bg-light small">${rows}</div>`);
        }
      } catch (_) { }
      restrictDocLinksIn(els.noDetailsBody);
    }
    noDetailsModal.show();
  }

  // Notice wiring
  if (els.noAddBtn) els.noAddBtn.addEventListener('click', noOnAdd);
  if (els.noForm) els.noForm.addEventListener('submit', noOnSubmit);
  if (els.noSearchInput) els.noSearchInput.addEventListener('input', debounce(() => { noCurrentPage = 1; renderNo(); }, 200));
  if (els.noTbody) { els.noTbody.addEventListener('click', (ev) => { const eBtn = ev.target.closest('.no-edit'); if (eBtn) { ev.preventDefault(); noOpenEdit(eBtn.getAttribute('data-id')); return; } const dBtn = ev.target.closest('.no-del'); if (dBtn) { ev.preventDefault(); noDelete(dBtn.getAttribute('data-id')); return; } const tr = ev.target.closest('tr'); if (!tr) return; const cell = ev.target.closest('td'); if (cell === tr.lastElementChild) return; const id = tr.getAttribute('data-id'); noOpenDetails(id); try { const it = noLoad().find(x => x.id === id) || {}; const body = els.noDetailsBody; if (body) { body.querySelector('#noMeetDateRow')?.remove(); body.querySelector('#noMeetTimeRow')?.remove(); const dateCol = `<div class="col-md-4" id="noMeetDateRow"><div class="small text-muted">Date of Meeting</div><div class="fw-semibold">${fmtDate(it.meetDate) || ''}</div></div>`; const timeCol = it.meetTime ? `<div class="col-md-4" id="noMeetTimeRow"><div class="small text-muted">Time of Meeting</div><div class="fw-semibold">${escapeHtml(fmtTime(it.meetTime))}</div></div>` : ''; const insertHtml = dateCol + timeCol; const labels = Array.from(body.querySelectorAll('.small.text-muted')); const subjLabel = labels.find(el => (el.textContent || '').trim() === 'Subject'); if (subjLabel && subjLabel.parentElement) { subjLabel.parentElement.insertAdjacentHTML('beforebegin', insertHtml); } else { const docDateLabel = labels.find(el => (el.textContent || '').trim() === 'Document Date'); if (docDateLabel && docDateLabel.parentElement) { docDateLabel.parentElement.insertAdjacentHTML('afterend', insertHtml); } else { body.insertAdjacentHTML('afterbegin', `<div class="row g-2">${insertHtml}</div>`); } } restrictDocLinksIn(body); } } catch (_) { } }); }
  // Global wrappers
  window.noEdit = (id) => { try { noOpenEdit(id); } catch (_) { } };
  window.noDelete = (id) => { try { noDelete(id); } catch (_) { } };
  window.noOpenDetails = (id) => { try { noOpenDetails(id); } catch (_) { } };

  // Enforce restricted access and inject Venue on Notice details after rendering
  try {
    const __origNoOpenDetails = noOpenDetails;
    noOpenDetails = function (id) {
      try { __origNoOpenDetails(id); } catch (_) { }
      try {
        restrictDocLinksIn(els.noDetailsBody);
        const body = els.noDetailsBody;
        const it = (noLoad().find(x => x.id === id) || {});
        if (body) {
          try { body.querySelector('#noVenueRow')?.remove(); } catch (_) { }
          const v = (it.venue || '').trim();
          if (v) {
            const venueCol = `<div class="col-md-4" id="noVenueRow"><div class="small text-muted">Venue</div><div class="fw-semibold">${escapeHtml(v)}</div></div>`;
            const labels = Array.from(body.querySelectorAll('.small.text-muted'));
            const signatoryLabel = labels.find(el => (el.textContent || '').trim() === 'Signatory');
            if (signatoryLabel && signatoryLabel.parentElement) {
              signatoryLabel.parentElement.insertAdjacentHTML('afterend', venueCol);
            } else {
              body.insertAdjacentHTML('afterbegin', `<div class="row g-2">${venueCol}</div>`);
            }
          }
        }
      } catch (_) { }
    };
  } catch (_) { }

  // ---- Office Order module ----
  const OO_PAGE_SIZE = 60;
  let ooCurrentPage = 1;
  let ooMultiDates = [];

  function ooGetMode() {
    if (els.ooModeMulti?.checked) return 'multi';
    if (els.ooModeSingle?.checked) return 'single';
    return 'range';
  }
  function ooSetMode(mode) {
    try {
      if (mode === 'multi') { els.ooModeMulti.checked = true; }
      else if (mode === 'single') { els.ooModeSingle.checked = true; }
      else { els.ooModeRange.checked = true; mode = 'range'; }
    } catch (_) { }
    // Show/hide wrappers
    if (els.ooRangeWrap) els.ooRangeWrap.classList.toggle('d-none', mode !== 'range');
    if (els.ooSingleWrap) els.ooSingleWrap.classList.toggle('d-none', mode !== 'single');
    if (els.ooMultiWrap) els.ooMultiWrap.classList.toggle('d-none', mode !== 'multi');
    // Required/disabled flags
    if (els.ooIncStart) { els.ooIncStart.required = (mode === 'range'); els.ooIncStart.disabled = (mode !== 'range'); }
    if (els.ooIncEnd) { els.ooIncEnd.required = (mode === 'range' && !els.ooNoEnd?.checked); els.ooIncEnd.disabled = (mode !== 'range' || !!els.ooNoEnd?.checked); }
    if (els.ooNoEnd) { els.ooNoEnd.disabled = (mode !== 'range'); }
    if (els.ooSingleDate) { els.ooSingleDate.required = (mode === 'single'); els.ooSingleDate.disabled = (mode !== 'single'); }
    if (els.ooMultiDateInput) { els.ooMultiDateInput.disabled = (mode !== 'multi'); }
  }

  function ooRenderMultiList() {
    if (!els.ooMultiList) return;
    const chips = (ooMultiDates || []).map(d =>
      `<span class="badge rounded-pill text-bg-light border fw-normal">
        ${fmtDate(d)}
        <button type="button" class="btn-close ms-1 align-middle oo-multi-remove" aria-label="Remove" data-date="${d}"></button>
      </span>`
    ).join(' ');
    els.ooMultiList.innerHTML = chips || '<span class="text-muted small">No dates added yet.</span>';
  }

  function ooAddMultiDate(val) {
    const d = normalizeYmd(val);
    if (!d) return;
    if (!ooMultiDates.includes(d)) ooMultiDates.push(d);
    ooMultiDates.sort();
    ooRenderMultiList();
  }
  function ooRemoveMultiDate(val) {
    const d = normalizeYmd(val);
    ooMultiDates = (ooMultiDates || []).filter(x => x !== d);
    ooRenderMultiList();
  }

  function ooRowHtml(it) {
    let range = '';
    const mode = it.incMode || 'range';
    if (mode === 'multi') {
      const arr = Array.isArray(it.incDates) ? it.incDates : [];
      range = arr.map(fmtDate).join(', ');
    } else if (mode === 'single') {
      const d = (Array.isArray(it.incDates) && it.incDates[0]) || it.incStart || '';
      range = fmtDate(d);
    } else {
      const start = fmtDate(it.incStart) || '';
      const end = fmtDate(it.incEnd) || '';
      range = (it.noEnd || !end) ? `${start} to Until revoked` : `${start} to ${end}`;
    }
    const delBtn = isAdmin ? `<button class=\"btn btn-sm btn-outline-danger oo-del\" data-id=\"${it.id}\" onclick=\"event.stopPropagation(); window.ooDelete && window.ooDelete('${it.id}')\"><i class=\"fa fa-trash\"></i></button>` : '';
    const editBtn = canEditDocReg ? `<button class=\"btn btn-sm btn-outline-primary me-1 oo-edit\" data-id=\"${it.id}\" onclick=\"event.stopPropagation(); window.ooEdit && window.ooEdit('${it.id}')\"><i class=\"fa fa-pen\"></i></button>` : '';
    return `<tr data-id="${it.id}" class="oo-row" onclick="window.ooOpenDetails && window.ooOpenDetails(this.getAttribute('data-id'))">
      <td>${escapeHtml(it.refNo || '')}</td>
      <td>${fmtDate(it.docDate) || ''}</td>
      <td>${escapeHtml(it.subject || '')}</td>
      <td>${escapeHtml(range)}</td>
      <td class="text-nowrap">${editBtn}${delBtn}</td>
    </tr>`;
  }

  function ooBuildPagination(total) {
    const ul = document.getElementById('ooPagination'); if (!ul) return;
    const maxPage = Math.max(1, Math.ceil(total / OO_PAGE_SIZE));
    if (ooCurrentPage > maxPage) ooCurrentPage = maxPage;
    let html = '';
    html += `<li class="page-item${ooCurrentPage <= 1 ? ' disabled' : ''}"><a class="page-link" href="#" data-page="prev">Prev</a></li>`;
    for (let p = 1; p <= maxPage; p++) html += `<li class="page-item${p === ooCurrentPage ? ' active' : ''}"><a class="page-link" href="#" data-page="${p}">${p}</a></li>`;
    html += `<li class="page-item${ooCurrentPage >= maxPage ? ' disabled' : ''}"><a class="page-link" href="#" data-page="next">Next</a></li>`;
    ul.innerHTML = html;
    ul.onclick = (ev) => {
      const a = ev.target.closest('a[data-page]'); if (!a) return; ev.preventDefault();
      const val = a.getAttribute('data-page');
      if (val === 'prev' && ooCurrentPage > 1) ooCurrentPage--;
      else if (val === 'next') { const max = Math.max(1, Math.ceil(total / OO_PAGE_SIZE)); if (ooCurrentPage < max) ooCurrentPage++; }
      else ooCurrentPage = parseInt(val, 10) || 1;
      renderOo();
    };
  }

  function renderOo() {
    const items = ooLoad().slice();
    const text = (els.ooSearchInput?.value || '').toLowerCase().trim();
    let list = items;
    if (text) {
      list = list.filter(it => {
        const hay = [it.refNo, it.subject, it.docLink, (it.participants || '')]
          .map(v => (v || '').toLowerCase()).join(' \u0000 ');
        return hay.includes(text);
      });
    }
    list.sort((a, b) => { const ad = a.docDate || '', bd = b.docDate || ''; if (ad !== bd) return ad < bd ? 1 : -1; const ac = a.createdAt || '', bc = b.createdAt || ''; return ac < bc ? 1 : -1; });
    const total = list.length;
    const start = (ooCurrentPage - 1) * OO_PAGE_SIZE;
    const end = Math.min(start + OO_PAGE_SIZE, total);
    const slice = list.slice(start, end);
    if (els.ooTbody) els.ooTbody.innerHTML = slice.map(ooRowHtml).join('');
    if (els.ooEmpty) els.ooEmpty.style.display = total ? 'none' : 'block';
    ooBuildPagination(total);
  }

  function ooSyncNoEnd() {
    const mode = (typeof ooGetMode === 'function') ? ooGetMode() : 'range';
    if (mode !== 'range') {
      // In single/multi modes, range inputs should not participate in validation
      if (els.ooIncStart) els.ooIncStart.required = false;
      if (els.ooIncEnd) { els.ooIncEnd.disabled = true; els.ooIncEnd.required = false; }
      return;
    }
    const noEnd = !!els.ooNoEnd?.checked;
    if (els.ooIncEnd) {
      els.ooIncEnd.disabled = noEnd;
      els.ooIncEnd.required = !noEnd;
      if (noEnd) els.ooIncEnd.value = '';
    }
  }
  function ooResetForm() {
    els.ooForm?.reset();
    try { if (els.ooDocDate) els.ooDocDate.value = new Date().toISOString().slice(0, 10); } catch (_) { }
    if (els.ooNoEnd) { els.ooNoEnd.checked = false; }
    ooSyncNoEnd();
    // reset modes
    ooSetMode('range');
    ooMultiDates = [];
    if (els.ooSingleDate) els.ooSingleDate.value = '';
    if (els.ooMultiDateInput) els.ooMultiDateInput.value = '';
    ooRenderMultiList();
  }
  function ooOnAdd() { if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; } ooResetForm(); if (els.ooId) els.ooId.value = ''; ooModal.show(); }
  function ooOnSubmit(e) {
    e.preventDefault(); if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    const mode = ooGetMode();
    // Enforce correct required/disabled flags based on current mode before validation
    ooSetMode(mode);
    ooSyncNoEnd();
    // Validation per mode
    if (mode === 'range') {
      const s = normalizeYmd(els.ooIncStart?.value || '');
      const eVal = normalizeYmd(els.ooIncEnd?.value || '');
      if (!s) { els.ooIncStart?.focus(); els.ooForm.reportValidity(); return; }
      if (!(els.ooNoEnd?.checked) && !eVal) { els.ooIncEnd?.focus(); els.ooForm.reportValidity(); return; }
    } else if (mode === 'single') {
      const sd = normalizeYmd(els.ooSingleDate?.value || '');
      if (!sd) { els.ooSingleDate?.focus(); els.ooForm.reportValidity(); return; }
    } else if (mode === 'multi') {
      if (!ooMultiDates.length) { alert('Please add at least one date.'); els.ooMultiDateInput?.focus(); return; }
    }
    if (!els.ooForm.checkValidity()) { els.ooForm.reportValidity(); return; }
    let incStart = '', incEnd = '', noEnd = false, incDates = [], incMode = mode;
    if (mode === 'range') {
      incStart = normalizeYmd(els.ooIncStart?.value || '');
      noEnd = !!els.ooNoEnd?.checked;
      incEnd = noEnd ? '' : normalizeYmd(els.ooIncEnd?.value || '');
      incDates = [];
    } else if (mode === 'single') {
      const sd = normalizeYmd(els.ooSingleDate?.value || '');
      incStart = sd; incEnd = sd; noEnd = false; incDates = [sd];
    } else if (mode === 'multi') {
      incDates = ooMultiDates.slice();
      incStart = ''; incEnd = ''; noEnd = false;
    }
    const it = {
      id: els.ooId?.value || ('id_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36)),
      refNo: (els.ooRefNo?.value || '').trim(),
      docDate: els.ooDocDate?.value || '',
      subject: (els.ooSubject?.value || '').trim(),
      incMode,
      incStart,
      incEnd,
      noEnd,
      incDates,
      participants: (els.ooParticipants?.value || '').trim(),
      docLink: (els.ooDocLink?.value || '').trim(),
      createdAt: els.ooId?.value ? undefined : new Date().toISOString(),
    };
    const items = ooLoad();
    if (els.ooId?.value && items.some(x => x.id === els.ooId.value)) { const idx = items.findIndex(x => x.id === it.id); if (idx >= 0) items[idx] = { ...items[idx], ...it }; }
    else items.push(it);
    ooSave(items);
    if (els.ooId) els.ooId.value = '';
    ooModal.hide(); renderOo();
  }
  function ooOpenEdit(id) {
    if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    const it = ooLoad().find(x => x.id === id); if (!it) return; ooResetForm();
    if (els.ooId) els.ooId.value = it.id;
    if (els.ooRefNo) els.ooRefNo.value = it.refNo || '';
    if (els.ooDocDate) els.ooDocDate.value = it.docDate || '';
    if (els.ooSubject) els.ooSubject.value = it.subject || '';
    const mode = it.incMode || (it.incStart && it.incEnd && it.incStart === it.incEnd ? 'single' : 'range');
    ooSetMode(mode);
    if (mode === 'range') {
      if (els.ooIncStart) els.ooIncStart.value = it.incStart || '';
      if (els.ooNoEnd) { els.ooNoEnd.checked = !!(it.noEnd || !it.incEnd); }
      ooSyncNoEnd();
      if (els.ooIncEnd && !els.ooNoEnd?.checked) els.ooIncEnd.value = it.incEnd || '';
    } else if (mode === 'single') {
      if (els.ooSingleDate) { const sd = (Array.isArray(it.incDates) && it.incDates[0]) || it.incStart || ''; els.ooSingleDate.value = sd; }
    } else if (mode === 'multi') {
      ooMultiDates = Array.isArray(it.incDates) ? it.incDates.slice() : [];
      ooRenderMultiList();
    }
    if (els.ooParticipants) els.ooParticipants.value = it.participants || '';
    if (els.ooDocLink) els.ooDocLink.value = it.docLink || '';
    ooModal.show();
  }
  function ooDelete(id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } const items = ooLoad(); const it = items.find(x => x.id === id); if (!it) return; if (!confirm(`Delete office order ${it.refNo || ''}? This cannot be undone.`)) return; ooSave(items.filter(x => x.id !== id)); renderOo(); }
  function ooOpenDetails(id) {
    const it = ooLoad().find(x => x.id === id); if (!it) return;
    let range = '';
    const mode = it.incMode || 'range';
    if (mode === 'multi') {
      const arr = Array.isArray(it.incDates) ? it.incDates : [];
      range = arr.map(fmtDate).join(', ');
    } else if (mode === 'single') {
      const d = (Array.isArray(it.incDates) && it.incDates[0]) || it.incStart || '';
      range = fmtDate(d);
    } else {
      const start = fmtDate(it.incStart) || '';
      const end = fmtDate(it.incEnd) || '';
      range = (it.noEnd || !end) ? (start + ' to Until revoked') : (start + ' to ' + end);
    }
    const participantsHtml = (it.participants || '').split(/\r?\n/).filter(Boolean).map(n => `<div>${escapeHtml(n)}</div>`).join('');
    const html = `
      <div class="row g-2">
        <div class="col-md-4"><div class="small text-muted">Reference #</div><div class="fw-semibold">${escapeHtml(it.refNo || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Document Date</div><div class="fw-semibold">${fmtDate(it.docDate) || ''}</div></div>
        <div class="col-md-4"><div class="small text-muted">Date Inclusion</div><div class="fw-semibold">${escapeHtml(range)}</div></div>
        <div class="col-12"><div class="small text-muted">Subject</div><div class="fw-semibold">${escapeHtml(it.subject || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Participants</div><div class="fw-semibold">${participantsHtml || ''}</div></div>
        <div class="col-12"><div class="small text-muted">Document Link</div><div class="fw-semibold">${renderDocLink(it.docLink)}</div></div>
      </div>
    `;
    if (els.ooDetailsBody) {
      els.ooDetailsBody.innerHTML = html;
      restrictDocLinksIn(els.ooDetailsBody);
      try {
        const __hist = (it.history || []).slice().sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        if (__hist.length) {
          const rows = __hist.map(h => {
            const who = (h.fullName || '').trim() || (h.email || 'unknown');
            const when = (new Date(h.timestamp)).toLocaleString();
            const act = h.action || 'update';
            return `<div class=\"d-flex align-items-start gap-2 py-1\"><i class=\"fa-regular fa-clock mt-1 text-muted\"></i><div><div><strong>${escapeHtml(act)}</strong> by ${escapeHtml(who)}</div><div class=\"text-muted small\">${escapeHtml(when)}</div></div></div>`;
          }).join('');
          els.ooDetailsBody.insertAdjacentHTML('beforeend', `<div class=\"small text-muted mt-3\">Edit History</div><div class=\"border rounded p-2 bg-light small\">${rows}</div>`);
        }
      } catch (_) { }
    }
    ooDetailsModal.show();
  }
  // Office Order wiring
  if (els.ooAddBtn) els.ooAddBtn.addEventListener('click', ooOnAdd);
  if (els.ooForm) els.ooForm.addEventListener('submit', ooOnSubmit);
  if (els.ooNoEnd) els.ooNoEnd.addEventListener('change', ooSyncNoEnd);
  if (els.ooModeRange) els.ooModeRange.addEventListener('change', () => ooSetMode('range'));
  if (els.ooModeSingle) els.ooModeSingle.addEventListener('change', () => ooSetMode('single'));
  if (els.ooModeMulti) els.ooModeMulti.addEventListener('change', () => ooSetMode('multi'));
  if (els.ooMultiAddBtn) els.ooMultiAddBtn.addEventListener('click', () => { const v = els.ooMultiDateInput?.value || ''; ooAddMultiDate(v); if (els.ooMultiDateInput) els.ooMultiDateInput.value = ''; });
  if (els.ooMultiDateInput) els.ooMultiDateInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); const v = els.ooMultiDateInput.value || ''; ooAddMultiDate(v); els.ooMultiDateInput.value = ''; } });
  if (els.ooMultiList) { els.ooMultiList.addEventListener('click', (ev) => { const b = ev.target.closest('.oo-multi-remove'); if (!b) return; ooRemoveMultiDate(b.getAttribute('data-date')); }); }
  if (els.ooSearchInput) els.ooSearchInput.addEventListener('input', debounce(() => { ooCurrentPage = 1; renderOo(); }, 200));
  if (els.ooTbody) {
    els.ooTbody.addEventListener('click', (ev) => {
      const eBtn = ev.target.closest('.oo-edit'); if (eBtn) { ev.preventDefault(); ooOpenEdit(eBtn.getAttribute('data-id')); return; }
      const dBtn = ev.target.closest('.oo-del'); if (dBtn) { ev.preventDefault(); ooDelete(dBtn.getAttribute('data-id')); return; }
      const tr = ev.target.closest('tr'); if (!tr) return; const cell = ev.target.closest('td'); if (cell === tr.lastElementChild) return; ooOpenDetails(tr.getAttribute('data-id'));
    });
  }
  // Global wrappers
  window.ooEdit = (id) => { try { ooOpenEdit(id); } catch (_) { } };
  window.ooDelete = (id) => { try { ooDelete(id); } catch (_) { } };
  window.ooOpenDetails = (id) => { try { ooOpenDetails(id); } catch (_) { } };

  // ---- Travel Order module ----
  const TO_PAGE_SIZE = 60;
  let toCurrentPage = 1;
  let toMultiDates = [];

  function toGetMode() {
    if (els.toModeMulti?.checked) return 'multi';
    if (els.toModeSingle?.checked) return 'single';
    return 'range';
  }
  function toSetMode(mode) {
    try {
      if (mode === 'multi') { els.toModeMulti.checked = true; }
      else if (mode === 'single') { els.toModeSingle.checked = true; }
      else { els.toModeRange.checked = true; mode = 'range'; }
    } catch (_) { }
    if (els.toRangeWrap) els.toRangeWrap.classList.toggle('d-none', mode !== 'range');
    if (els.toSingleWrap) els.toSingleWrap.classList.toggle('d-none', mode !== 'single');
    if (els.toMultiWrap) els.toMultiWrap.classList.toggle('d-none', mode !== 'multi');
    // toggle required/disabled
    if (els.toIncStart) { els.toIncStart.required = (mode === 'range'); els.toIncStart.disabled = (mode !== 'range'); }
    if (els.toIncEnd) { els.toIncEnd.required = (mode === 'range' && !els.toNoEnd?.checked); els.toIncEnd.disabled = (mode !== 'range' || !!els.toNoEnd?.checked); }
    if (els.toNoEnd) { els.toNoEnd.disabled = (mode !== 'range'); }
    if (els.toSingleDate) { els.toSingleDate.required = (mode === 'single'); els.toSingleDate.disabled = (mode !== 'single'); }
    if (els.toMultiDateInput) { els.toMultiDateInput.disabled = (mode !== 'multi'); }
  }

  function toRenderMultiList() {
    if (!els.toMultiList) return;
    const chips = (toMultiDates || []).map(d =>
      `<span class="badge rounded-pill text-bg-light border fw-normal">${fmtDate(d)}<button type="button" class="btn-close ms-1 align-middle to-multi-remove" aria-label="Remove" data-date="${d}"></button></span>`
    ).join(' ');
    els.toMultiList.innerHTML = chips || '<span class="text-muted small">No dates added yet.</span>';
  }
  function toAddMultiDate(val) { const d = normalizeYmd(val); if (!d) return; if (!toMultiDates.includes(d)) toMultiDates.push(d); toMultiDates.sort(); toRenderMultiList(); }
  function toRemoveMultiDate(val) { const d = normalizeYmd(val); toMultiDates = (toMultiDates || []).filter(x => x !== d); toRenderMultiList(); }

  function toRowHtml(it) {
    let range = '';
    const mode = it.incMode || 'range';
    if (mode === 'multi') {
      const arr = Array.isArray(it.incDates) ? it.incDates : [];
      range = arr.map(fmtDate).join(', ');
    } else if (mode === 'single') {
      const d = (Array.isArray(it.incDates) && it.incDates[0]) || it.incStart || '';
      range = fmtDate(d);
    } else {
      const start = fmtDate(it.incStart) || '';
      const end = fmtDate(it.incEnd) || '';
      range = (it.noEnd || !end) ? `${start} to Until revoked` : `${start} to ${end}`;
    }
    const delBtn = isAdmin ? `<button class=\"btn btn-sm btn-outline-danger to-del\" data-id=\"${it.id}\" onclick=\"event.stopPropagation(); window.toDelete && window.toDelete('${it.id}')\"><i class=\"fa fa-trash\"></i></button>` : '';
    const editBtn = canEditDocReg ? `<button class=\"btn btn-sm btn-outline-primary me-1 to-edit\" data-id=\"${it.id}\" onclick=\"event.stopPropagation(); window.toEdit && window.toEdit('${it.id}')\"><i class=\"fa fa-pen\"></i></button>` : '';
    return `<tr data-id="${it.id}" class="to-row" onclick="window.toOpenDetails && window.toOpenDetails(this.getAttribute('data-id'))">
      <td>${escapeHtml(it.refNo || '')}</td>
      <td>${fmtDate(it.docDate) || ''}</td>
      <td>${escapeHtml(it.subject || '')}</td>
      <td>${escapeHtml(range)}</td>
      <td class="text-nowrap">${editBtn}${delBtn}</td>
    </tr>`;
  }

  function toBuildPagination(total) {
    const ul = document.getElementById('toPagination'); if (!ul) return;
    const maxPage = Math.max(1, Math.ceil(total / TO_PAGE_SIZE));
    if (toCurrentPage > maxPage) toCurrentPage = maxPage;
    let html = '';
    html += `<li class="page-item${toCurrentPage <= 1 ? ' disabled' : ''}"><a class="page-link" href="#" data-page="prev">Prev</a></li>`;
    for (let p = 1; p <= maxPage; p++) html += `<li class="page-item${p === toCurrentPage ? ' active' : ''}"><a class="page-link" href="#" data-page="${p}">${p}</a></li>`;
    html += `<li class="page-item${toCurrentPage >= maxPage ? ' disabled' : ''}"><a class="page-link" href="#" data-page="next">Next</a></li>`;
    ul.innerHTML = html;
    ul.onclick = (ev) => {
      const a = ev.target.closest('a[data-page]'); if (!a) return; ev.preventDefault();
      const val = a.getAttribute('data-page');
      if (val === 'prev' && toCurrentPage > 1) toCurrentPage--;
      else if (val === 'next') { const max = Math.max(1, Math.ceil(total / TO_PAGE_SIZE)); if (toCurrentPage < max) toCurrentPage++; }
      else toCurrentPage = parseInt(val, 10) || 1;
      renderTo();
    };
  }

  function renderTo() {
    const items = toLoad().slice();
    const text = (els.toSearchInput?.value || '').toLowerCase().trim();
    let list = items;
    if (text) {
      list = list.filter(it => {
        const hay = [it.refNo, it.subject, it.docLink, (it.participants || '')]
          .map(v => (v || '').toLowerCase()).join(' \u0000 ');
        return hay.includes(text);
      });
    }
    list.sort((a, b) => { const ad = a.docDate || '', bd = b.docDate || ''; if (ad !== bd) return ad < bd ? 1 : -1; const ac = a.createdAt || '', bc = b.createdAt || ''; return ac < bc ? 1 : -1; });
    const total = list.length;
    const start = (toCurrentPage - 1) * TO_PAGE_SIZE;
    const end = Math.min(start + TO_PAGE_SIZE, total);
    const slice = list.slice(start, end);
    if (els.toTbody) els.toTbody.innerHTML = slice.map(toRowHtml).join('');
    if (els.toEmpty) els.toEmpty.style.display = total ? 'none' : 'block';
    toBuildPagination(total);
  }

  function toSyncNoEnd() {
    const mode = (typeof toGetMode === 'function') ? toGetMode() : 'range';
    if (mode !== 'range') {
      if (els.toIncStart) els.toIncStart.required = false;
      if (els.toIncEnd) { els.toIncEnd.disabled = true; els.toIncEnd.required = false; }
      return;
    }
    const noEnd = !!els.toNoEnd?.checked;
    if (els.toIncEnd) {
      els.toIncEnd.disabled = noEnd;
      els.toIncEnd.required = !noEnd;
      if (noEnd) els.toIncEnd.value = '';
    }
  }
  function toResetForm() {
    els.toForm?.reset();
    try { if (els.toDocDate) els.toDocDate.value = new Date().toISOString().slice(0, 10); } catch (_) { }
    if (els.toNoEnd) { els.toNoEnd.checked = false; }
    toSetMode('range');
    toMultiDates = [];
    if (els.toSingleDate) els.toSingleDate.value = '';
    if (els.toMultiDateInput) els.toMultiDateInput.value = '';
    toRenderMultiList();
    toSyncNoEnd();
  }
  function toOnAdd() { if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; } toResetForm(); if (els.toId) els.toId.value = ''; toModal.show(); }
  function toOnSubmit(e) {
    e.preventDefault(); if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    const mode = toGetMode();
    toSetMode(mode);
    toSyncNoEnd();
    // per-mode validation
    if (mode === 'range') {
      const s = normalizeYmd(els.toIncStart?.value || '');
      const eVal = normalizeYmd(els.toIncEnd?.value || '');
      if (!s) { els.toIncStart?.focus(); els.toForm.reportValidity(); return; }
      if (!(els.toNoEnd?.checked) && !eVal) { els.toIncEnd?.focus(); els.toForm.reportValidity(); return; }
    } else if (mode === 'single') {
      const sd = normalizeYmd(els.toSingleDate?.value || '');
      if (!sd) { els.toSingleDate?.focus(); els.toForm.reportValidity(); return; }
    } else if (mode === 'multi') {
      if (!toMultiDates.length) { alert('Please add at least one date.'); els.toMultiDateInput?.focus(); return; }
    }
    if (!els.toForm.checkValidity()) { els.toForm.reportValidity(); return; }
    let incStart = '', incEnd = '', noEnd = false, incDates = [], incMode = mode;
    if (mode === 'range') {
      incStart = normalizeYmd(els.toIncStart?.value || '');
      noEnd = !!els.toNoEnd?.checked;
      incEnd = noEnd ? '' : normalizeYmd(els.toIncEnd?.value || '');
    } else if (mode === 'single') {
      const sd = normalizeYmd(els.toSingleDate?.value || '');
      incStart = sd; incEnd = sd; incDates = [sd];
    } else if (mode === 'multi') {
      incDates = toMultiDates.slice();
    }
    const it = {
      id: els.toId?.value || uid(),
      refNo: (els.toRefNo?.value || '').trim(),
      docDate: els.toDocDate?.value || '',
      subject: (els.toSubject?.value || '').trim(),
      incMode,
      incStart,
      incEnd,
      noEnd,
      incDates,
      participants: (els.toParticipants?.value || '').trim(),
      docLink: (els.toDocLink?.value || '').trim(),
      createdAt: els.toId?.value ? undefined : new Date().toISOString(),
    };
    const items = toLoad();
    if (els.toId?.value && items.some(x => x.id === els.toId.value)) { const idx = items.findIndex(x => x.id === it.id); if (idx >= 0) items[idx] = { ...items[idx], ...it }; }
    else items.push(it);
    toSave(items);
    if (els.toId) els.toId.value = '';
    toModal.hide(); renderTo();
  }
  function toOpenEdit(id) {
    if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    const it = toLoad().find(x => x.id === id); if (!it) return; toResetForm();
    if (els.toId) els.toId.value = it.id;
    if (els.toRefNo) els.toRefNo.value = it.refNo || '';
    if (els.toDocDate) els.toDocDate.value = it.docDate || '';
    if (els.toSubject) els.toSubject.value = it.subject || '';
    const mode = it.incMode || (it.incStart && it.incEnd && it.incStart === it.incEnd ? 'single' : 'range');
    toSetMode(mode);
    if (mode === 'range') {
      if (els.toIncStart) els.toIncStart.value = it.incStart || '';
      if (els.toNoEnd) { els.toNoEnd.checked = !!(it.noEnd || !it.incEnd); }
      toSyncNoEnd();
      if (els.toIncEnd && !els.toNoEnd?.checked) els.toIncEnd.value = it.incEnd || '';
    } else if (mode === 'single') {
      if (els.toSingleDate) { const sd = (Array.isArray(it.incDates) && it.incDates[0]) || it.incStart || ''; els.toSingleDate.value = sd; }
    } else if (mode === 'multi') {
      toMultiDates = Array.isArray(it.incDates) ? it.incDates.slice() : [];
      toRenderMultiList();
    }
    if (els.toParticipants) els.toParticipants.value = it.participants || '';
    if (els.toDocLink) els.toDocLink.value = it.docLink || '';
    toModal.show();
  }
  function toDelete(id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } const items = toLoad(); const it = items.find(x => x.id === id); if (!it) return; if (!confirm(`Delete travel order ${it.refNo || ''}? This cannot be undone.`)) return; toSave(items.filter(x => x.id !== id)); renderTo(); }
  function toOpenDetails(id) {
    const it = toLoad().find(x => x.id === id); if (!it) return;
    let range = '';
    const mode = it.incMode || 'range';
    if (mode === 'multi') {
      const arr = Array.isArray(it.incDates) ? it.incDates : []; range = arr.map(fmtDate).join(', ');
    } else if (mode === 'single') {
      const d = (Array.isArray(it.incDates) && it.incDates[0]) || it.incStart || ''; range = fmtDate(d);
    } else {
      const start = fmtDate(it.incStart) || ''; const end = fmtDate(it.incEnd) || ''; range = (it.noEnd || !end) ? (start + ' to Until revoked') : (start + ' to ' + end);
    }
    const participantsHtml = (it.participants || '').split(/\r?\n/).filter(Boolean).map(n => `<div>${escapeHtml(n)}</div>`).join('');
    const html = `
      <div class="row g-2">
        <div class="col-md-4"><div class="small text-muted">Reference #</div><div class="fw-semibold">${escapeHtml(it.refNo || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Document Date</div><div class="fw-semibold">${fmtDate(it.docDate) || ''}</div></div>
        <div class="col-md-4"><div class="small text-muted">Date Inclusion</div><div class="fw-semibold">${escapeHtml(range)}</div></div>
        <div class="col-12"><div class="small text-muted">Subject</div><div class="fw-semibold">${escapeHtml(it.subject || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Participants</div><div class="fw-semibold">${participantsHtml || ''}</div></div>
        <div class="col-12"><div class="small text-muted">Document Link</div><div class="fw-semibold">${renderDocLink(it.docLink)}</div></div>
      </div>
    `;
    if (els.toDetailsBody) {
      els.toDetailsBody.innerHTML = html;
      restrictDocLinksIn(els.toDetailsBody);
      try {
        const __hist = (it.history || []).slice().sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        if (__hist.length) {
          const rows = __hist.map(h => {
            const who = (h.fullName || '').trim() || (h.email || 'unknown');
            const when = (new Date(h.timestamp)).toLocaleString();
            const act = h.action || 'update';
            return `<div class=\"d-flex align-items-start gap-2 py-1\"><i class=\"fa-regular fa-clock mt-1 text-muted\"></i><div><div><strong>${escapeHtml(act)}</strong> by ${escapeHtml(who)}</div><div class=\"text-muted small\">${escapeHtml(when)}</div></div></div>`;
          }).join('');
          els.toDetailsBody.insertAdjacentHTML('beforeend', `<div class=\"small text-muted mt-3\">Edit History</div><div class=\"border rounded p-2 bg-light small\">${rows}</div>`);
        }
      } catch (_) { }
      els.toDetailsBody.setAttribute('data-current-id', it.id);
    }
    // Ensure footer has Edit/Delete
    try {
      const footer = els.toDetailsModalEl?.querySelector('.modal-footer');
      if (footer) {
        const editBtnHtml = canEditDocReg ? `<button type=\"button\" class=\"btn btn-primary me-auto\" id=\"toEditBtn\">Edit</button>` : '';
        const delBtnHtml = isAdmin ? `<button type="button" class="btn btn-danger" id="toDelBtn">Delete</button>` : '';
        footer.innerHTML = `${editBtnHtml}
          ${delBtnHtml}
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>`;
        const editBtnEl = footer.querySelector('#toEditBtn');
        if (editBtnEl) { editBtnEl.addEventListener('click', () => { toDetailsModal.hide(); toOpenEdit(it.id); }); }
        const delBtnEl = footer.querySelector('#toDelBtn');
        if (delBtnEl) { delBtnEl.addEventListener('click', () => { toDelete(it.id); }); }
      }
    } catch (_) { }
    toDetailsModal.show();
  }
  // Travel wiring
  if (els.toAddBtn) els.toAddBtn.addEventListener('click', toOnAdd);
  if (els.toForm) els.toForm.addEventListener('submit', toOnSubmit);
  if (els.toNoEnd) els.toNoEnd.addEventListener('change', toSyncNoEnd);
  if (els.toModeRange) els.toModeRange.addEventListener('change', () => toSetMode('range'));
  if (els.toModeSingle) els.toModeSingle.addEventListener('change', () => toSetMode('single'));
  if (els.toModeMulti) els.toModeMulti.addEventListener('change', () => toSetMode('multi'));
  if (els.toMultiAddBtn) els.toMultiAddBtn.addEventListener('click', () => { const v = els.toMultiDateInput?.value || ''; toAddMultiDate(v); if (els.toMultiDateInput) els.toMultiDateInput.value = ''; });
  if (els.toMultiDateInput) els.toMultiDateInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); const v = els.toMultiDateInput.value || ''; toAddMultiDate(v); els.toMultiDateInput.value = ''; } });
  if (els.toMultiList) { els.toMultiList.addEventListener('click', (ev) => { const b = ev.target.closest('.to-multi-remove'); if (!b) return; toRemoveMultiDate(b.getAttribute('data-date')); }); }
  if (els.toSearchInput) els.toSearchInput.addEventListener('input', debounce(() => { toCurrentPage = 1; renderTo(); }, 200));
  if (els.toTbody) {
    els.toTbody.addEventListener('click', (ev) => {
      const eBtn = ev.target.closest('.to-edit'); if (eBtn) { ev.preventDefault(); toOpenEdit(eBtn.getAttribute('data-id')); return; }
      const dBtn = ev.target.closest('.to-del'); if (dBtn) { ev.preventDefault(); toDelete(dBtn.getAttribute('data-id')); return; }
      const tr = ev.target.closest('tr'); if (!tr) return; const cell = ev.target.closest('td'); if (cell === tr.lastElementChild) return; toOpenDetails(tr.getAttribute('data-id'));
    });
  }
  // Global wrappers
  window.toEdit = (id) => { try { toOpenEdit(id); } catch (_) { } };
  window.toDelete = (id) => { try { toDelete(id); } catch (_) { } };
  window.toOpenDetails = (id) => { try { toOpenDetails(id); } catch (_) { } };

  // ---- Minutes of Meeting module ----
  const MI_PAGE_SIZE = 60;
  let miCurrentPage = 1;

  function miLoad() { try { const raw = localStorage.getItem(MI_KEY); const arr = raw ? JSON.parse(raw) : []; return Array.isArray(arr) ? arr : []; } catch (_) { return []; } }
  function miSave(items) { try { localStorage.setItem(MI_KEY, JSON.stringify(items || [])); } catch (_) { } }

  function miRowHtml(it) {
    const editBtn = canEditDocReg ? `<button class="btn btn-sm btn-outline-primary me-1 mi-edit" data-id="${it.id}" onclick="event.stopPropagation(); window.miEdit && window.miEdit('${it.id}')"><i class="fa fa-pen"></i></button>` : '';
    return `<tr data-id="${it.id}" class="mi-row" onclick="window.miOpenDetails && window.miOpenDetails(this.getAttribute('data-id'))">
      <td>${escapeHtml(it.controlNo || '')}</td>
      <td>${fmtDate(it.docDate) || ''}</td>
      <td>${fmtDate(it.meetDate) || ''}</td>
      <td>${escapeHtml(it.subject || '')}</td>
      <td>${escapeHtml(it.signatory || '')}</td>
      <td class="text-nowrap">
        ${editBtn}
        ${isAdmin ? `<button class=\"btn btn-sm btn-outline-danger mi-del\" data-id=\"${it.id}\" onclick=\"event.stopPropagation(); window.miDelete && window.miDelete('${it.id}')\"><i class=\"fa fa-trash\"></i></button>` : ''}
      </td>
    </tr>`;
  }

  function miBuildPagination(total) {
    const ul = document.getElementById('miPagination'); if (!ul) return;
    const maxPage = Math.max(1, Math.ceil(total / MI_PAGE_SIZE));
    if (miCurrentPage > maxPage) miCurrentPage = maxPage;
    let html = '';
    html += `<li class="page-item${miCurrentPage <= 1 ? ' disabled' : ''}"><a class="page-link" href="#" data-page="prev">Prev</a></li>`;
    for (let p = 1; p <= maxPage; p++) html += `<li class="page-item${p === miCurrentPage ? ' active' : ''}"><a class="page-link" href="#" data-page="${p}">${p}</a></li>`;
    html += `<li class="page-item${miCurrentPage >= maxPage ? ' disabled' : ''}"><a class="page-link" href="#" data-page="next">Next</a></li>`;
    ul.innerHTML = html;
    ul.onclick = (ev) => {
      const a = ev.target.closest('a[data-page]'); if (!a) return; ev.preventDefault();
      const val = a.getAttribute('data-page');
      if (val === 'prev' && miCurrentPage > 1) miCurrentPage--;
      else if (val === 'next') { const max = Math.max(1, Math.ceil(total / MI_PAGE_SIZE)); if (miCurrentPage < max) miCurrentPage++; }
      else miCurrentPage = parseInt(val, 10) || 1;
      renderMi();
    };
  }

  function renderMi() {
    const items = miLoad().slice();
    const text = (els.miSearchInput?.value || '').toLowerCase().trim();
    let list = items;
    if (text) {
      list = list.filter(it => {
        const hay = [it.controlNo, it.subject, it.signatory, it.preparedBy, it.docLink]
          .map(v => (v || '').toLowerCase()).join(' \u0000 ');
        return hay.includes(text);
      });
    }
    list.sort((a, b) => { const ad = a.docDate || '', bd = b.docDate || ''; if (ad !== bd) return ad < bd ? 1 : -1; const ac = a.createdAt || '', bc = b.createdAt || ''; return ac < bc ? 1 : -1; });
    const total = list.length;
    const start = (miCurrentPage - 1) * MI_PAGE_SIZE;
    const end = Math.min(start + MI_PAGE_SIZE, total);
    const slice = list.slice(start, end);
    if (els.miTbody) els.miTbody.innerHTML = slice.map(miRowHtml).join('');
    if (els.miEmpty) els.miEmpty.style.display = total ? 'none' : 'block';
    miBuildPagination(total);
  }

  function miResetForm() {
    els.miForm?.reset();
    try { if (els.miDocDate) els.miDocDate.value = new Date().toISOString().slice(0, 10); } catch (_) { }
  }
  function miOnAdd() { if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; } miResetForm(); if (els.miId) els.miId.value = ''; miModal.show(); }
  function miOnSubmit(e) {
    e.preventDefault(); if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    if (!els.miForm.checkValidity()) { els.miForm.reportValidity(); return; }
    const it = {
      id: els.miId?.value || uid(),
      controlNo: (els.miControlNo?.value || '').trim(),
      docDate: els.miDocDate?.value || '',
      meetDate: els.miMeetDate?.value || '',
      subject: (els.miSubject?.value || '').trim(),
      signatory: (els.miSignatory?.value || '').trim(),
      agenda: (els.miAgenda?.value || '').trim(),
      preparedBy: (els.miPreparedBy?.value || '').trim(),
      docLink: (els.miDocLink?.value || '').trim(),
      remarks: (els.miRemarks?.value || '').trim(),
      createdAt: els.miId?.value ? undefined : new Date().toISOString(),
    };
    const items = miLoad();
    if (els.miId?.value && items.some(x => x.id === els.miId.value)) { const idx = items.findIndex(x => x.id === it.id); if (idx >= 0) items[idx] = { ...items[idx], ...it }; }
    else items.push(it);
    miSave(items);
    if (els.miId) els.miId.value = '';
    miModal.hide(); renderMi();
  }
  function miOpenEdit(id) {
    if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    const it = miLoad().find(x => x.id === id); if (!it) return; miResetForm();
    if (els.miId) els.miId.value = it.id;
    if (els.miControlNo) els.miControlNo.value = it.controlNo || '';
    if (els.miDocDate) els.miDocDate.value = it.docDate || '';
    if (els.miMeetDate) els.miMeetDate.value = it.meetDate || '';
    if (els.miSubject) els.miSubject.value = it.subject || '';
    if (els.miSignatory) els.miSignatory.value = it.signatory || '';
    if (els.miAgenda) els.miAgenda.value = it.agenda || '';
    if (els.miPreparedBy) els.miPreparedBy.value = it.preparedBy || '';
    if (els.miDocLink) els.miDocLink.value = it.docLink || '';
    if (els.miRemarks) els.miRemarks.value = it.remarks || '';
    miModal.show();
  }
  function miDelete(id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } const items = miLoad(); const it = items.find(x => x.id === id); if (!it) return; if (!confirm(`Delete minutes ${it.controlNo || ''}? This cannot be undone.`)) return; miSave(items.filter(x => x.id !== id)); renderMi(); }
  function miOpenDetails(id) {
    const it = miLoad().find(x => x.id === id); if (!it) return;
    const agendaHtml = (it.agenda || '').split(/\r?\n/).filter(Boolean).map(n => `<div>${escapeHtml(n)}</div>`).join('');
    const html = `
      <div class="row g-2">
        <div class="col-md-3"><div class="small text-muted">Control No.</div><div class="fw-semibold">${escapeHtml(it.controlNo || '')}</div></div>
        <div class="col-md-3"><div class="small text-muted">Document Date</div><div class="fw-semibold">${fmtDate(it.docDate) || ''}</div></div>
        <div class="col-md-3"><div class="small text-muted">Date of Meeting</div><div class="fw-semibold">${fmtDate(it.meetDate) || ''}</div></div>
        <div class="col-md-3"><div class="small text-muted">Signatory</div><div class="fw-semibold">${escapeHtml(it.signatory || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Subject</div><div class="fw-semibold">${escapeHtml(it.subject || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Agenda</div><div class="fw-semibold">${agendaHtml || ''}</div></div>
        <div class="col-md-6"><div class="small text-muted">Prepared by</div><div class="fw-semibold">${escapeHtml(it.preparedBy || '')}</div></div>
        <div class="col-md-6"><div class="small text-muted">Document Link</div><div class="fw-semibold">${renderDocLink(it.docLink)}</div></div>
        <div class="col-12"><div class="small text-muted">Remarks</div><div class="fw-semibold">${escapeHtml(it.remarks || '')}</div></div>
      </div>
    `;
    if (els.miDetailsBody) {
      els.miDetailsBody.innerHTML = html;
      restrictDocLinksIn(els.miDetailsBody);
      try {
        const __hist = (it.history || []).slice().sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        if (__hist.length) {
          const rows = __hist.map(h => {
            const who = (h.fullName || '').trim() || (h.email || 'unknown');
            const when = (new Date(h.timestamp)).toLocaleString();
            const act = h.action || 'update';
            return `<div class=\"d-flex align-items-start gap-2 py-1\"><i class=\"fa-regular fa-clock mt-1 text-muted\"></i><div><div><strong>${escapeHtml(act)}</strong> by ${escapeHtml(who)}</div><div class=\"text-muted small\">${escapeHtml(when)}</div></div></div>`;
          }).join('');
          els.miDetailsBody.insertAdjacentHTML('beforeend', `<div class=\"small text-muted mt-3\">Edit History</div><div class=\"border rounded p-2 bg-light small\">${rows}</div>`);
        }
      } catch (_) { }
    }
    miDetailsModal.show();
  }
  // Minutes wiring
  if (els.miAddBtn) els.miAddBtn.addEventListener('click', miOnAdd);
  if (els.miForm) els.miForm.addEventListener('submit', miOnSubmit);
  if (els.miSearchInput) els.miSearchInput.addEventListener('input', debounce(() => { miCurrentPage = 1; renderMi(); }, 200));
  if (els.miTbody) {
    els.miTbody.addEventListener('click', (ev) => {
      const eBtn = ev.target.closest('.mi-edit'); if (eBtn) { ev.preventDefault(); miOpenEdit(eBtn.getAttribute('data-id')); return; }
      const dBtn = ev.target.closest('.mi-del'); if (dBtn) { ev.preventDefault(); miDelete(dBtn.getAttribute('data-id')); return; }
      const tr = ev.target.closest('tr'); if (!tr) return; const cell = ev.target.closest('td'); if (cell === tr.lastElementChild) return; miOpenDetails(tr.getAttribute('data-id'));
    });
  }
  // Global wrappers
  window.miEdit = (id) => { try { miOpenEdit(id); } catch (_) { } };
  window.miDelete = (id) => { try { miDelete(id); } catch (_) { } };
  window.miOpenDetails = (id) => { try { miOpenDetails(id); } catch (_) { } };

  // ---- Incoming Letters module (parallel to Outgoing) ----
  const IL_KEY = 'docreg_incoming_letters_v1';
  const IL_PAGE_SIZE = 60;
  let ilCurrentPage = 1;

  function ilLoad() { try { const raw = localStorage.getItem(IL_KEY); const arr = raw ? JSON.parse(raw) : []; return Array.isArray(arr) ? arr : []; } catch (_) { return []; } }
  function ilSave(items) { try { localStorage.setItem(IL_KEY, JSON.stringify(items || [])); } catch (_) { } }

  function ilRowHtml(it) {
    const name = (it.fromName || it.name || it.toName || '');
    const designation = (it.fromDesignation || it.designation || it.toDesignation || '');
    const agency = (it.fromAgency || it.agency || it.toAgency || '');
    const address = (it.fromAddress || it.address || it.toAddress || '');
    const addrBlock = `
      <div class="fw-semibold">${escapeHtml(name)}</div>
      <div class="text-muted small">${escapeHtml(designation)}</div>
      <div class="text-muted small">${escapeHtml(agency)}</div>
      <div class="text-muted small">${escapeHtml(address)}</div>
    `;
    const delBtn = isAdmin ? `<button class=\"btn btn-sm btn-outline-danger il-del\" data-id=\"${it.id}\" onclick=\"event.stopPropagation(); window.ilDelete && window.ilDelete('${it.id}')\"><i class=\"fa fa-trash\"></i></button>` : '';
    const editBtn = canEditDocReg ? `<button class="btn btn-sm btn-outline-primary me-1 il-edit" data-id="${it.id}" onclick="event.stopPropagation(); window.ilEdit && window.ilEdit('${it.id}')"><i class="fa fa-pen"></i></button>` : '';
    return `<tr data-id="${it.id}" class="il-row" onclick="window.ilOpenDetails && window.ilOpenDetails(this.getAttribute('data-id'))">
      <td>${escapeHtml(it.refNo || '')}</td>
      <td>${fmtDate(it.docDate) || ''}</td>
      <td>${addrBlock}</td>
      <td>${escapeHtml(it.subject || '')}</td>
      <td>${fmtDate(it.receivedDate) || ''}</td>
      <td class="text-nowrap">
        ${editBtn}
        ${delBtn}
      </td>
    </tr>`;
  }

  function buildIlPagination(total) {
    const ul = document.getElementById('ilPagination'); if (!ul) return;
    const maxPage = Math.max(1, Math.ceil(total / IL_PAGE_SIZE));
    if (ilCurrentPage > maxPage) ilCurrentPage = maxPage;
    let html = '';
    html += `<li class="page-item${ilCurrentPage <= 1 ? ' disabled' : ''}"><a class="page-link" href="#" data-page="prev">Prev</a></li>`;
    for (let p = 1; p <= maxPage; p++) html += `<li class="page-item${p === ilCurrentPage ? ' active' : ''}"><a class="page-link" href="#" data-page="${p}">${p}</a></li>`;
    html += `<li class="page-item${ilCurrentPage >= maxPage ? ' disabled' : ''}"><a class="page-link" href="#" data-page="next">Next</a></li>`;
    ul.innerHTML = html;
    ul.onclick = (ev) => {
      const a = ev.target.closest('a[data-page]'); if (!a) return; ev.preventDefault();
      const val = a.getAttribute('data-page');
      if (val === 'prev' && ilCurrentPage > 1) ilCurrentPage--;
      else if (val === 'next') { const max = Math.max(1, Math.ceil(total / IL_PAGE_SIZE)); if (ilCurrentPage < max) ilCurrentPage++; }
      else ilCurrentPage = parseInt(val, 10) || 1;
      renderIl();
    };
  }

  function renderIl() {
    const items = ilLoad().slice();
    // Build delivery methods
    try {
      const uniq = Array.from(new Set(items.map(i => (i.deliveryVia || '').trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b));
      if (els.ilDeliveryFilter) {
        const curr = els.ilDeliveryFilter.value || '';
        els.ilDeliveryFilter.innerHTML = ['<option value="">All Methods</option>'].concat(uniq.map(v => `<option${v === curr ? ' selected' : ''}>${escapeHtml(v)}</option>`)).join('');
      }
    } catch (_) { }
    // Filters
    const text = (els.ilSearchInput?.value || '').toLowerCase().trim();
    const status = els.ilStatusFilter?.value || '';
    const via = els.ilDeliveryFilter?.value || '';
    let list = items;
    if (status) list = list.filter(it => (it.status || '') === status);
    if (via) list = list.filter(it => (it.deliveryVia || '') === via);
    if (text) {
      list = list.filter(it => {
        const hay = [
          it.refNo,
          it.name, it.designation, it.agency, it.address,
          it.toName, it.toDesignation, it.toAgency, it.toAddress,
          it.fromName, it.fromDesignation, it.fromAgency, it.fromAddress,
          it.subject, it.preparedBy, it.category, it.assignedTo
        ].map(v => (v || '').toLowerCase()).join(' \u0000 ');
        return hay.includes(text);
      });
    }
    // Sort latest first
    list.sort((a, b) => { const ad = a.docDate || '', bd = b.docDate || ''; if (ad !== bd) return ad < bd ? 1 : -1; const ac = a.createdAt || '', bc = b.createdAt || ''; return ac < bc ? 1 : -1; });
    const total = list.length;
    const start = (ilCurrentPage - 1) * IL_PAGE_SIZE;
    const end = Math.min(start + IL_PAGE_SIZE, total);
    const slice = list.slice(start, end);
    if (els.ilTbody) els.ilTbody.innerHTML = slice.map(ilRowHtml).join('');
    if (els.ilEmpty) els.ilEmpty.style.display = total ? 'none' : 'block';
    buildIlPagination(total);
  }

  function ilResetForm() { els.ilForm?.reset(); }
  function ilOnAdd() { if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; } ilResetForm(); if (els.ilId) els.ilId.value = ''; try { const today = new Date().toISOString().slice(0, 10); els.ilDocDate.value = today; if (els.ilReceivedDate) els.ilReceivedDate.value = today; } catch (_) { } ilModal.show(); }
  function ilOnSubmit(e) {
    e.preventDefault(); if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; } if (!els.ilForm.checkValidity()) { els.ilForm.reportValidity(); return; }
    // Gather structured TO/FROM; fallback to legacy party block if present
    const toName = (els.ilToName?.value || '').trim();
    const toDesignation = (els.ilToDesignation?.value || '').trim();
    const toAgency = (els.ilToAgency?.value || '').trim();
    const toAddress = (els.ilToAddress?.value || '').trim();
    const fromName = (els.ilFromName?.value || '').trim();
    const fromPosition = (els.ilFromPosition?.value || '').trim();
    const fromAgency = (els.ilFromAgency?.value || '').trim();
    const fromAddress = (els.ilFromAddress?.value || '').trim();
    const blockLines = (els.ilPartyBlock?.value || '').split(/\r?\n/).map(s => s.trim());
    const legacyName = blockLines[0] || '';
    const legacyDesignation = blockLines[1] || '';
    const legacyAgency = blockLines[2] || '';
    const legacyAddress = blockLines[3] || '';
    const it = {
      id: els.ilId?.value || uid(),
      refNo: els.ilRefNo.value.trim(),
      docDate: els.ilDocDate.value || '',
      // Legacy display block maps to TO by default
      name: toName || legacyName,
      designation: toDesignation || legacyDesignation,
      agency: toAgency || legacyAgency,
      address: toAddress || legacyAddress,
      toName, toDesignation, toAgency, toAddress,
      fromName, fromPosition, fromAgency, fromAddress,
      subject: els.ilSubject.value.trim(),
      category: (els.ilCategory?.value || '').trim(),
      receivedDate: (els.ilReceivedDate?.value || '').trim(),
      docLink: (els.ilDocLink?.value || '').trim(),
      deliveryVia: els.ilDeliveryVia.value.trim(),
      assignedTo: (els.ilAssignedTo?.value || '').trim(),
      status: els.ilStatus.value,
      remarks: els.ilRemarks.value.trim(),
      createdAt: els.ilId?.value ? undefined : new Date().toISOString(),
    };
    const items = ilLoad();
    if (els.ilId?.value) { const idx = items.findIndex(x => x.id === it.id); if (idx >= 0) items[idx] = { ...items[idx], ...it }; }
    else items.push(it);
    ilSave(items); ilModal.hide(); renderIl();
  }
  function ilOpenEdit(id) {
    if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    const it = ilLoad().find(x => x.id === id); if (!it) return; ilResetForm();
    if (els.ilId) els.ilId.value = it.id;
    els.ilRefNo.value = it.refNo || '';
    els.ilDocDate.value = it.docDate || '';
    if (els.ilToName) els.ilToName.value = it.toName || it.name || '';
    if (els.ilToDesignation) els.ilToDesignation.value = it.toDesignation || it.designation || '';
    if (els.ilToAgency) els.ilToAgency.value = it.toAgency || it.agency || '';
    if (els.ilToAddress) els.ilToAddress.value = it.toAddress || it.address || '';
    if (els.ilFromName) els.ilFromName.value = it.fromName || '';
    if (els.ilFromPosition) els.ilFromPosition.value = it.fromPosition || it.fromDesignation || '';
    if (els.ilFromAgency) els.ilFromAgency.value = it.fromAgency || '';
    if (els.ilFromAddress) els.ilFromAddress.value = it.fromAddress || '';
    if (els.ilPartyBlock) { const lines = [it.name || '', it.designation || '', it.agency || '', it.address || '']; els.ilPartyBlock.value = lines.join('\n').replace(/\n+$/, ''); }
    els.ilSubject.value = it.subject || '';
    if (els.ilDocLink) els.ilDocLink.value = it.docLink || '';
    if (els.ilCategory) els.ilCategory.value = it.category || '';
    if (els.ilReceivedDate) els.ilReceivedDate.value = (it.receivedDate || '');
    els.ilDeliveryVia.value = it.deliveryVia || '';
    if (els.ilAssignedTo) els.ilAssignedTo.value = it.assignedTo || '';
    els.ilStatus.value = it.status || '';
    els.ilRemarks.value = it.remarks || '';
    ilModal.show();
  }
  function ilDelete(id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } const items = ilLoad(); const it = items.find(x => x.id === id); if (!it) return; if (!confirm(`Delete document ${it.refNo || ''}? This cannot be undone.`)) return; ilSave(items.filter(x => x.id !== id)); renderIl(); }
  function ilOpenDetails(id) {
    console.log('[docreg] ilOpenDetails called with id:', id);
    const allItems = ilLoad();
    console.log('[docreg] ilLoad() returned', allItems.length, 'items');
    const it = allItems.find(x => x.id === id);
    console.log('[docreg] Found item:', it);
    if (!it) { console.log('[docreg] Item not found, returning early'); return; }
    // compute reply metrics
    const outs = loadItems();
    const memos = moLoad();
    const notices = noLoad();
    const docMap = new Map(outs.concat(memos).concat(notices).map(o => [o.id, o]));
    const unlinkedSet = new Set(Array.isArray(it.unlinkedReplyIds) ? it.unlinkedReplyIds : []);
    // Start from stored replies, if any
    let replies = Array.isArray(it.replies) ? it.replies.slice() : [];
    // Also consider memoranda that explicitly link back to this incoming (in case linkage write failed)
    try {
      const memLinked = memos
        .filter(m => (m && m.replyToIncomingId === it.id) && !unlinkedSet.has(m.id))
        .map(m => ({ id: m.id, type: 'memo', date: (m.docDate || '') }));
      if (memLinked.length) {
        const map = new Map();
        for (const r of replies) { if (r && r.id) map.set(r.id, { id: r.id, type: (r.type || 'outgoing'), date: (r.date || '') }); }
        for (const r of memLinked) { if (r && r.id) map.set(r.id, { id: r.id, type: r.type, date: (r.date || '') }); }
        const merged = Array.from(map.values());
        if (merged.length !== replies.length) {
          replies = merged;
          // Persist the healed replies list for this incoming
          try { const db = window.firebase?.firestore?.(); if (db) { db.collection('docreg_incoming').doc(it.id).set({ replies }, { merge: true }); } } catch (_) { }
        }
      }
    } catch (_) { }
    // Prune replies to only those explicitly linked or with exact ref/control match; also drop deleted docs
    try {
      const anchor = (it.refNo || '').trim();
      const norm = (s) => (s || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '').trim();
      const anchorNorm = norm(anchor);
      const filtered = replies.filter(r => {
        if (!r || !r.id) return false;
        const o = docMap.get(r.id);
        if (!o) return false;
        if (o.replyToIncomingId === it.id) return true;
        if (anchorNorm) {
          const a = norm(o.refNo), b = norm(o.controlNo);
          if ((a && a === anchorNorm) || (b && b === anchorNorm)) return true;
        }
        return false;
      });
      if (filtered.length !== replies.length) {
        replies = filtered;
        try { const db = window.firebase?.firestore?.(); if (db) { db.collection('docreg_incoming').doc(it.id).set({ replies }, { merge: true }); } else { const arr = ilLoad(); const idx = arr.findIndex(x => x.id === it.id); if (idx >= 0) { arr[idx] = { ...arr[idx], replies }; ilSave(arr); } } } catch (_) { }
      }
    } catch (_) { }
    // Anchor heuristic: include docs that reference this Reference No. by exact refNo/controlNo match only
    try {
      const anchor = (it.refNo || '').trim();
      if (anchor) {
        const norm = (s) => (s || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '').trim();
        const anchorNorm = norm(anchor);
        const map = new Map();
        for (const r of replies) { if (r && r.id) { map.set(r.id, { id: r.id, type: (r.type || 'outgoing'), date: (r.date || '') }); } }
        for (const o of outs) {
          const refMatch = (() => { const a = norm(o.refNo), b = norm(o.controlNo); return (a && a === anchorNorm) || (b && b === anchorNorm); })();
          if (refMatch && !map.has(o.id) && !unlinkedSet.has(o.id)) { map.set(o.id, { id: o.id, type: 'outgoing', date: (o.docDate || '') }); }
        }
        for (const m of memos) {
          const refMatch = (() => { const a = norm(m.refNo), b = norm(m.controlNo); return (a && a === anchorNorm) || (b && b === anchorNorm); })();
          if (refMatch && !map.has(m.id) && !unlinkedSet.has(m.id)) { map.set(m.id, { id: m.id, type: 'memo', date: (m.docDate || '') }); }
        }
        for (const n of notices) {
          const refMatch = (() => { const a = norm(n.refNo), b = norm(n.controlNo); return (a && a === anchorNorm) || (b && b === anchorNorm); })();
          if (refMatch && !map.has(n.id) && !unlinkedSet.has(n.id)) { map.set(n.id, { id: n.id, type: 'notice', date: (n.docDate || '') }); }
        }
        const merged = Array.from(map.values());
        if (merged.length !== replies.length) {
          replies = merged;
          try { const db = window.firebase?.firestore?.(); if (db) { db.collection('docreg_incoming').doc(it.id).set({ replies }, { merge: true }); } else { const arr = ilLoad(); const idx = arr.findIndex(x => x.id === it.id); if (idx >= 0) { arr[idx] = { ...arr[idx], replies }; ilSave(arr); } } } catch (_) { }
        }
      }
    } catch (_) { }
    try {
      let changed = false;
      replies = replies.map(r => {
        if (!r || !r.id) return r;
        const doc = docMap.get(r.id);
        const docDate = doc?.docDate || '';
        if (docDate && docDate !== (r.date || '')) {
          changed = true;
          return { ...r, date: docDate };
        }
        return r;
      });
      if (changed) {
        try { const db = window.firebase?.firestore?.(); if (db) { db.collection('docreg_incoming').doc(it.id).set({ replies }, { merge: true }); } else { const arr = ilLoad(); const idx = arr.findIndex(x => x.id === it.id); if (idx >= 0) { arr[idx] = { ...arr[idx], replies }; ilSave(arr); } } } catch (_) { }
      }
    } catch (_) { }
    let firstReplyDays = '';
    try {
      if (replies.length) {
        const recvYmd = it.receivedDate || it.docDate || '';
        const replyYmds = replies.map(r => (r.date || docMap.get(r.id)?.docDate || '')).filter(Boolean);
        if (replyYmds.length && recvYmd) {
          // find earliest reply date
          const minYmd = replyYmds.reduce((min, cur) => {
            const ma = Date.parse((min || '') + 'T00:00:00');
            const mb = Date.parse((cur || '') + 'T00:00:00');
            if (isNaN(ma)) return cur; if (isNaN(mb)) return min; return mb < ma ? cur : min;
          }, replyYmds[0]);
          const days = businessDaysDiff(recvYmd, minYmd);
          if (isFinite(days)) firstReplyDays = `${days} working day${days === 1 ? '' : 's'}`;
        }
      }
    } catch (_) { }
    const repliesBlock = replies.length ? (
      '<div class="col-12"><div class="small text-muted">Action/Replies</div>' +
      replies.map(r => {
        const o = docMap.get(r.id) || {}; const dt = o.docDate || r.date || ''; const ref = (o.controlNo || o.refNo || r.id); const recv = it.receivedDate || it.docDate || '';
        let diff = '';
        try { const d = businessDaysDiff(recv, dt); if (isFinite(d)) diff = ` (${d} working day${d === 1 ? '' : 's'})`; } catch (_) { }
        const unlinkBtn = isAdmin ? `<button class="btn btn-sm btn-outline-danger ms-2 il-unlink-reply" data-reply-id="${escapeHtml(r.id)}" title="Unlink">Unlink</button>` : '';
        const rtype = (r.type || 'outgoing').toLowerCase();
        const rlabel = (rtype === 'notice') ? 'NOTICE' : (rtype === 'memo') ? 'MEMO' : 'OUTGOING';
        return `<div class="d-flex align-items-center justify-content-between"><div class="fw-semibold text-primary cursor-pointer text-decoration-underline-hover il-reply-link" data-id="${escapeHtml(r.id)}" data-type="${escapeHtml(rtype)}" title="Click to view details">${rlabel} - ${escapeHtml(ref)} on ${fmtDate(dt)}${diff}</div>${unlinkBtn}</div>`;
      }).join('') +
      (firstReplyDays ? `<div class="text-muted small">First reply in ${firstReplyDays}</div>` : '') +
      '</div>'
    ) : '';
    const fromName = (it.fromName || it.name || it.toName || '').trim();
    const fromDesignation = (it.fromDesignation || it.designation || it.toDesignation || '').trim();
    const fromAgency = (it.fromAgency || it.agency || it.toAgency || '').trim();
    const fromAddress = (it.fromAddress || it.address || it.toAddress || '').trim();
    const letterFromParts = [];
    if (fromName) letterFromParts.push(`<div class="fw-semibold">${escapeHtml(fromName)}</div>`);
    if (fromDesignation) letterFromParts.push(`<div>${escapeHtml(fromDesignation)}</div>`);
    if (fromAgency) letterFromParts.push(`<div>${escapeHtml(fromAgency)}</div>`);
    if (fromAddress) letterFromParts.push(`<div>${escapeHtml(fromAddress)}</div>`);
    const letterFromHtml = letterFromParts.join('');
    const html = `
      <div class="row g-2">
        <div class="col-md-4"><div class="small text-muted">Reference No.</div><div class="fw-semibold">${escapeHtml(it.refNo || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Document Date</div><div class="fw-semibold">${fmtDate(it.docDate) || ''}</div></div>
        <div class="col-md-4"><div class="small text-muted">Document Received Date</div><div class="fw-semibold">${fmtDate(it.receivedDate) || ''}</div></div>
        <div class="col-md-4"><div class="small text-muted">Document Category</div><div class="fw-semibold">${escapeHtml(it.category || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Letter from:</div><div>${letterFromHtml}</div></div>
        <div class="col-12"><div class="small text-muted">Subject</div><div class="fw-semibold">${escapeHtml(it.subject || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Document Link</div><div class="fw-semibold">${renderDocLink(it.docLink)}</div></div>
        <div class="col-md-6"><div class="small text-muted">Delivery/Sent via</div><div class="fw-semibold">${escapeHtml(it.deliveryVia || '')}</div></div>
        <div class="col-md-6"><div class="small text-muted">Assigned To</div><div class="fw-semibold">${escapeHtml(it.assignedTo || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Status</div><div class="fw-semibold">${escapeHtml(it.status || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Remarks</div><div class="fw-semibold">${escapeHtml(it.remarks || '')}</div></div>
        ${repliesBlock}
      </div>`;
    if (els.ilDetailsBody) {
      els.ilDetailsBody.innerHTML = html;
      restrictDocLinksIn(els.ilDetailsBody);
      try {
        const __hist = (it.history || []).slice().sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        if (__hist.length) {
          const rows = __hist.map(h => {
            const who = (h.fullName || '').trim() || (h.email || 'unknown');
            const when = (new Date(h.timestamp)).toLocaleString();
            const act = h.action || 'update';
            return `<div class=\"d-flex align-items-start gap-2 py-1\"><i class=\"fa-regular fa-clock mt-1 text-muted\"></i><div><div><strong>${escapeHtml(act)}</strong> by ${escapeHtml(who)}</div><div class=\"text-muted small\">${escapeHtml(when)}</div></div></div>`;
          }).join('');
          els.ilDetailsBody.insertAdjacentHTML('beforeend', `<div class=\"small text-muted mt-3\">Edit History</div><div class=\"border rounded p-2 bg-light small\">${rows}</div>`);
        }
      } catch (_) { }
      els.ilDetailsBody.setAttribute('data-current-id', it.id);
      try {
        els.ilDetailsBody.querySelectorAll('.il-unlink-reply').forEach(btn => {
          btn.addEventListener('click', (ev) => {
            ev.preventDefault(); ev.stopPropagation();
            const rid = btn.getAttribute('data-reply-id');
            if (!rid) return;
            if (!confirm('Unlink this reply from the incoming document?')) return;
            try { window.ilUnlinkReply && window.ilUnlinkReply(it.id, rid); } catch (_) { }
          });
        });
      } catch (_) { }
      try {
        els.ilDetailsBody.querySelectorAll('.il-reply-link').forEach(link => {
          link.addEventListener('click', (ev) => {
            ev.preventDefault();
            const rid = link.getAttribute('data-id');
            const rtype = link.getAttribute('data-type');
            if (!rid || !rtype) return;
            ilDetailsModal.hide();
            if (rtype === 'outgoing' || rtype === 'out') { if (window.olOpenDetails) window.olOpenDetails(rid); }
            else if (rtype === 'memo' || rtype === 'memorandum') { if (window.moOpenDetails) window.moOpenDetails(rid); }
            else if (rtype === 'notice') { if (window.noOpenDetails) window.noOpenDetails(rid); }
          });
        });
      } catch (_) { }
    }
    console.log('[docreg] About to show ilDetailsModal:', ilDetailsModal);
    ilDetailsModal.show();
    console.log('[docreg] ilDetailsModal.show() called');
  }

  // Reply flow
  function openReplyChoice(incomingId) { replyContext = { incomingId }; replyChoiceModal.show(); }
  function prefillOutgoingFromIncoming(incoming) {
    // Switch to Outgoing section and open the Outgoing modal prefilled
    switchSection('outgoing');
    resetForm();
    if (els.refNo) els.refNo.value = `${incoming.refNo || ''}`;
    els.docDate.value = new Date().toISOString().slice(0, 10);
    if (els.olPartyBlock) {
      const lines = [incoming.name || '', incoming.designation || '', incoming.agency || '', incoming.address || ''];
      els.olPartyBlock.value = lines.join('\n').replace(/\n+$/, '');
    }
    // Structured TO/FROM
    if (els.olToName) els.olToName.value = incoming.name || '';
    if (els.olToDesignation) els.olToDesignation.value = incoming.designation || '';
    if (els.olToAgency) els.olToAgency.value = incoming.agency || '';
    if (els.olToAddress) els.olToAddress.value = incoming.address || '';
    if (els.olFromName) els.olFromName.value = '';
    if (els.olFromDesignation) els.olFromDesignation.value = '';
    if (els.olFromAgency) els.olFromAgency.value = '';
    if (els.olFromAddress) els.olFromAddress.value = '';
    els.subject.value = `RE: ${incoming.subject || ''}`.trim();
    els.preparedBy.value = '';
    // Linkage: temporarily stash incomingId on the form element
    els.form.setAttribute('data-reply-to', incoming.id);
    olModal.show();
  }
  function prefillNoticeFromIncoming(incoming) {
    // Switch to Notice section and open the Notice modal prefilled
    switchSection('notice');
    noResetForm();
    if (els.noRefNo) els.noRefNo.value = `${incoming.refNo || ''}`;
    if (els.noControlNo) els.noControlNo.value = '';
    if (els.noDocDate) els.noDocDate.value = new Date().toISOString().slice(0, 10);
    if (els.noPartyBlock) {
      const lines = [incoming.name || '', incoming.agency || ''];
      els.noPartyBlock.value = lines.join('\n').replace(/\n+$/, '');
    }
    if (els.noSubject) els.noSubject.value = `RE: ${incoming.subject || ''}`.trim();
    if (els.noSignatory) els.noSignatory.value = '';
    els.noForm?.setAttribute('data-reply-to', incoming.id);
    noModal.show();
  }
  function finalizeReplyLinkage(createdOutgoing) {
    const incomingId = els.form.getAttribute('data-reply-to');
    if (!incomingId) return;
    els.form.removeAttribute('data-reply-to');
    // attach linkage on both sides in localStorage
    try {
      // Outgoing side already saved; ensure it carries replyToIncomingId
      const outItems = loadItems();
      const idx = outItems.findIndex(x => x.id === createdOutgoing.id);
      if (idx >= 0) { outItems[idx] = { ...outItems[idx], replyToIncomingId: incomingId }; saveItems(outItems); }
      // Incoming side store replies array
      const inItems = ilLoad();
      const i = inItems.findIndex(x => x.id === incomingId);
      if (i >= 0) { const r = inItems[i].replies || []; r.push({ id: createdOutgoing.id, type: 'outgoing', date: createdOutgoing.docDate }); inItems[i] = { ...inItems[i], replies: r }; ilSave(inItems); }
    } catch (_) { }
  }

  // Incoming wiring
  console.log('[docreg] Incoming wiring: ilTbody =', els.ilTbody);
  if (els.ilAddBtn) els.ilAddBtn.addEventListener('click', ilOnAdd);
  if (els.ilForm) els.ilForm.addEventListener('submit', ilOnSubmit);
  if (els.ilSearchInput) els.ilSearchInput.addEventListener('input', debounce(() => { ilCurrentPage = 1; renderIl(); }, 200));
  if (els.ilStatusFilter) els.ilStatusFilter.addEventListener('change', () => { ilCurrentPage = 1; renderIl(); });
  if (els.ilDeliveryFilter) els.ilDeliveryFilter.addEventListener('change', () => { ilCurrentPage = 1; renderIl(); });
  if (els.ilTbody) { els.ilTbody.addEventListener('click', (ev) => { console.log('[docreg] ilTbody clicked:', ev.target); const eBtn = ev.target.closest('.il-edit'); if (eBtn) { ev.preventDefault(); ilOpenEdit(eBtn.getAttribute('data-id')); return; } const dBtn = ev.target.closest('.il-del'); if (dBtn) { ev.preventDefault(); ilDelete(dBtn.getAttribute('data-id')); return; } const tr = ev.target.closest('tr'); console.log('[docreg] ilTbody tr:', tr); if (!tr) return; const cell = ev.target.closest('td'); if (cell === tr.lastElementChild) return; console.log('[docreg] calling ilOpenDetails for', tr.getAttribute('data-id')); ilOpenDetails(tr.getAttribute('data-id')); }); }

  // Global wrappers
  window.ilEdit = (id) => { try { ilOpenEdit(id); } catch (_) { } };
  window.ilDelete = (id) => { try { ilDelete(id); } catch (_) { } };
  window.ilOpenDetails = (id) => { console.log('[docreg] window.ilOpenDetails called with', id); try { ilOpenDetails(id); } catch (_) { } };

  // Wire reply UI
  const ilReplyBtn = document.getElementById('ilReplyBtn');
  if (ilReplyBtn) { ilReplyBtn.addEventListener('click', () => { replyContext = { incomingId: els.ilDetailsBody?.dataset.currentId || '' }; replyChoiceModal.show(); }); }
  const replyOutgoingBtn = document.getElementById('replyOutgoingBtn');
  if (replyOutgoingBtn) {
    replyOutgoingBtn.addEventListener('click', () => {
      try {
        // use currently shown details to find the item
        const items = ilLoad();
        const last = els.ilDetailsBody && els.ilDetailsBody.getAttribute('data-current-id');
        const it = items.find(x => x.id === (last || replyContext?.incomingId));
        if (!it) { replyChoiceModal.hide(); return; }
        replyChoiceModal.hide(); ilDetailsModal.hide(); prefillOutgoingFromIncoming(it);
      } catch (_) { replyChoiceModal.hide(); }
    });
  }
  const replyMemoBtn = document.getElementById('replyMemoBtn');
  if (replyMemoBtn) {
    replyMemoBtn.addEventListener('click', () => {
      try {
        const items = ilLoad();
        const last = els.ilDetailsBody && els.ilDetailsBody.getAttribute('data-current-id');
        const it = items.find(x => x.id === (last || replyContext?.incomingId));
        if (!it) { replyChoiceModal.hide(); return; }
        replyChoiceModal.hide(); ilDetailsModal.hide(); prefillMemoFromIncoming(it);
      } catch (_) { replyChoiceModal.hide(); }
    });
  }
  const replyNoticeBtn = document.getElementById('replyNoticeBtn');
  if (replyNoticeBtn) {
    replyNoticeBtn.addEventListener('click', () => {
      try {
        const items = ilLoad();
        const last = els.ilDetailsBody && els.ilDetailsBody.getAttribute('data-current-id');
        const it = items.find(x => x.id === (last || replyContext?.incomingId));
        if (!it) { replyChoiceModal.hide(); return; }
        replyChoiceModal.hide(); ilDetailsModal.hide(); prefillNoticeFromIncoming(it);
      } catch (_) { replyChoiceModal.hide(); }
    });
  }

  // ---- Memorandum module ----
  const MO_PAGE_SIZE = 60;
  let moCurrentPage = 1;

  function moRowHtml(it) {
    const addrBlock = `
      <div class="fw-semibold">${escapeHtml(it.name || '')}</div>
      <div class="text-muted small">${escapeHtml(it.designation || '')}</div>
      <div class="text-muted small">${escapeHtml(it.agency || '')}</div>
      <div class="text-muted small">${escapeHtml(it.address || '')}</div>
    `;
    const delBtn = isAdmin ? `<button class=\"btn btn-sm btn-outline-danger mo-del\" data-id=\"${it.id}\" onclick=\"event.stopPropagation(); window.moDelete && window.moDelete('${it.id}')\"><i class=\"fa fa-trash\"></i></button>` : '';
    const editBtn = canEditDocReg ? `<button class="btn btn-sm btn-outline-primary me-1 mo-edit" data-id="${it.id}" onclick="event.stopPropagation(); window.moEdit && window.moEdit('${it.id}')"><i class="fa fa-pen"></i></button>` : '';
    return `<tr data-id="${it.id}" class="mo-row" onclick="window.moOpenDetails && window.moOpenDetails(this.getAttribute('data-id'))">
      <td>${escapeHtml((it.controlNo || it.refNo) || '')}</td>
      <td>${fmtDate(it.docDate) || ''}</td>
      <td>${addrBlock}</td>
      <td>${escapeHtml(it.subject || '')}</td>
      <td>${escapeHtml(it.preparedBy || '')}</td>
      <td class="text-nowrap">
        ${editBtn}
        ${delBtn}
      </td>
    </tr>`;
  }

  function moBuildPagination(total) {
    const ul = document.getElementById('moPagination'); if (!ul) return;
    const maxPage = Math.max(1, Math.ceil(total / MO_PAGE_SIZE));
    if (moCurrentPage > maxPage) moCurrentPage = maxPage;
    let html = '';
    html += `<li class="page-item${moCurrentPage <= 1 ? ' disabled' : ''}"><a class="page-link" href="#" data-page="prev">Prev</a></li>`;
    for (let p = 1; p <= maxPage; p++) html += `<li class="page-item${p === moCurrentPage ? ' active' : ''}"><a class="page-link" href="#" data-page="${p}">${p}</a></li>`;
    html += `<li class="page-item${moCurrentPage >= maxPage ? ' disabled' : ''}"><a class="page-link" href="#" data-page="next">Next</a></li>`;
    ul.innerHTML = html;
    ul.onclick = (ev) => {
      const a = ev.target.closest('a[data-page]'); if (!a) return; ev.preventDefault();
      const val = a.getAttribute('data-page');
      if (val === 'prev' && moCurrentPage > 1) moCurrentPage--;
      else if (val === 'next') { const max = Math.max(1, Math.ceil(total / MO_PAGE_SIZE)); if (moCurrentPage < max) moCurrentPage++; }
      else moCurrentPage = parseInt(val, 10) || 1;
      renderMo();
    };
  }

  function renderMo() {
    const items = moLoad().slice();
    const text = (els.moSearchInput?.value || '').toLowerCase().trim();
    const status = els.moStatusFilter?.value || '';
    let list = items;
    if (status) list = list.filter(it => (it.status || '') === status);
    if (text) {
      list = list.filter(it => {
        const hay = [it.controlNo, it.refNo, it.name, it.designation, it.agency, it.address, it.subject, it.preparedBy, it.signatory]
          .map(v => (v || '').toLowerCase()).join(' \u0000 ');
        return hay.includes(text);
      });
    }
    list.sort((a, b) => { const ad = a.docDate || '', bd = b.docDate || ''; if (ad !== bd) return ad < bd ? 1 : -1; const ac = a.createdAt || '', bc = b.createdAt || ''; return ac < bc ? 1 : -1; });
    const total = list.length;
    const start = (moCurrentPage - 1) * MO_PAGE_SIZE;
    const end = Math.min(start + MO_PAGE_SIZE, total);
    const slice = list.slice(start, end);
    if (els.moTbody) els.moTbody.innerHTML = slice.map(moRowHtml).join('');
    if (els.moEmpty) els.moEmpty.style.display = total ? 'none' : 'block';
    moBuildPagination(total);
  }

  function moResetForm() { els.moForm?.reset(); }
  function moOnAdd() { if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; } moResetForm(); try { if (els.moDocDate) els.moDocDate.value = new Date().toISOString().slice(0, 10); } catch (_) { } moModal.show(); }
  function moOnSubmit(e) {
    e.preventDefault(); if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; } if (!els.moForm.checkValidity()) { els.moForm.reportValidity(); return; }
    const partyLines = (els.moPartyBlock?.value || '').split(/\r?\n/).map(s => s.trim());
    const toName = (els.moToName?.value || '').trim() || partyLines[0] || '';
    const toDesignation = (els.moToDesignation?.value || '').trim() || partyLines[1] || '';
    const toDepartment = (els.moToDepartment?.value || '').trim() || partyLines[2] || '';
    const fromName = (els.moFromName?.value || '').trim() || '';
    const fromDesignation = (els.moFromDesignation?.value || '').trim() || '';
    const fromDepartment = (els.moFromDepartment?.value || '').trim() || '';
    const it = {
      id: els.moId?.value || ('id_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36)),
      refNo: (els.moRefNo?.value || '').trim(),
      controlNo: (els.moControlNo?.value || '').trim(),
      docDate: els.moDocDate?.value || '',
      // Legacy fields map to TO
      name: toName,
      designation: toDesignation,
      agency: toDepartment,
      address: partyLines[3] || '',
      // New structured fields
      toName,
      toDesignation,
      toDepartment,
      fromName,
      fromDesignation,
      fromDepartment,
      subject: (els.moSubject?.value || '').trim(),
      preparedBy: (els.moPreparedBy?.value || '').trim(),
      docLink: (els.moDocLink?.value || '').trim(),
      signatory: (els.moSignatory?.value || '').trim(),
      status: els.moStatus?.value || '',
      remarks: (els.moRemarks?.value || '').trim(),
      createdAt: els.moId?.value ? undefined : new Date().toISOString(),
      replyToIncomingId: els.moForm?.getAttribute('data-reply-to') || undefined,
    };
    const items = moLoad();
    if (els.moId?.value) { const idx = items.findIndex(x => x.id === it.id); if (idx >= 0) items[idx] = { ...items[idx], ...it }; }
    else items.push(it);
    moSave(items);
    try { moFinalizeReplyLinkage(it); } catch (_) { }
    moModal.hide(); renderMo();
  }
  function moOpenEdit(id) {
    if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
    const it = moLoad().find(x => x.id === id); if (!it) return; moResetForm();
    if (els.moId) els.moId.value = it.id;
    if (els.moRefNo) els.moRefNo.value = it.refNo || '';
    if (els.moControlNo) els.moControlNo.value = it.controlNo || '';
    if (els.moDocDate) els.moDocDate.value = it.docDate || '';
    // Populate new structured fields with stored values or fallback to legacy
    if (els.moToName) els.moToName.value = (it.toName || it.name || '');
    if (els.moToDesignation) els.moToDesignation.value = (it.toDesignation || it.designation || '');
    if (els.moToDepartment) els.moToDepartment.value = (it.toDepartment || it.agency || '');
    if (els.moFromName) els.moFromName.value = (it.fromName || '');
    if (els.moFromDesignation) els.moFromDesignation.value = (it.fromDesignation || '');
    if (els.moFromDepartment) els.moFromDepartment.value = (it.fromDepartment || '');
    // Back-compat textarea if present
    if (els.moPartyBlock) {
      const lines = [it.name || '', it.designation || '', it.agency || '', it.address || ''];
      els.moPartyBlock.value = lines.join('\n').replace(/\n+$/, '');
    }
    if (els.moSubject) els.moSubject.value = it.subject || '';
    if (els.moDocLink) els.moDocLink.value = it.docLink || '';
    if (els.moPreparedBy) els.moPreparedBy.value = it.preparedBy || '';
    if (els.moSignatory) els.moSignatory.value = it.signatory || '';
    if (els.moStatus) els.moStatus.value = it.status || '';
    if (els.moRemarks) els.moRemarks.value = it.remarks || '';
    moModal.show();
  }
  function moDelete(id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } const items = moLoad(); const it = items.find(x => x.id === id); if (!it) return; if (!confirm(`Delete memorandum ${it.refNo || it.controlNo || ''}? This cannot be undone.`)) return; moSave(items.filter(x => x.id !== id)); renderMo(); }
  function moOpenDetails(id) {
    const it = moLoad().find(x => x.id === id); if (!it) return;
    const toName = it.toName || it.name || '';
    const toDesignation = it.toDesignation || it.designation || '';
    const toDepartment = it.toDepartment || it.agency || '';
    const fromName = it.fromName || '';
    const fromDesignation = it.fromDesignation || '';
    const fromDepartment = it.fromDepartment || '';
    const memoToParts = [];
    if (toName) memoToParts.push(`<div class="fw-semibold">${escapeHtml(toName)}</div>`);
    if (toDesignation) memoToParts.push(`<div>${escapeHtml(toDesignation)}</div>`);
    if (toDepartment) memoToParts.push(`<div>${escapeHtml(toDepartment)}</div>`);
    const memoToHtml = memoToParts.join('');
    const memoFromParts = [];
    if (fromName) memoFromParts.push(`<div class="fw-semibold">${escapeHtml(fromName)}</div>`);
    if (fromDesignation) memoFromParts.push(`<div>${escapeHtml(fromDesignation)}</div>`);
    if (fromDepartment) memoFromParts.push(`<div>${escapeHtml(fromDepartment)}</div>`);
    const memoFromHtml = memoFromParts.join('');
    const html = `
      <div class="row g-2">
        <div class="col-md-4"><div class="small text-muted">Reference No.</div><div class="fw-semibold">${escapeHtml(it.refNo || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Document Date</div><div class="fw-semibold">${fmtDate(it.docDate) || ''}</div></div>
        <div class="col-md-4"><div class="small text-muted">Prepared by</div><div class="fw-semibold">${escapeHtml(it.preparedBy || '')}</div></div>
        <div class="col-md-4"><div class="small text-muted">Control No.</div><div class="fw-semibold">${escapeHtml(it.controlNo || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Memo to:</div><div>${memoToHtml}</div></div>
        <div class="col-12"><div class="small text-muted">Memo from:</div><div>${memoFromHtml}</div></div>
        <div class="col-12"><div class="small text-muted">Subject</div><div class="fw-semibold">${escapeHtml(it.subject || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Document Link</div><div class="fw-semibold">${renderDocLink(it.docLink)}</div></div>
        <div class="col-md-6"><div class="small text-muted">Signatory</div><div class="fw-semibold">${escapeHtml(it.signatory || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Status</div><div class="fw-semibold">${escapeHtml(it.status || '')}</div></div>
        <div class="col-12"><div class="small text-muted">Remarks</div><div class="fw-semibold">${escapeHtml(it.remarks || '')}</div></div>
      </div>
    `;
    if (els.moDetailsBody) {
      els.moDetailsBody.innerHTML = html;
      restrictDocLinksIn(els.moDetailsBody);
      try { els.moDetailsBody.setAttribute('data-current-id', it.id); } catch (_) { }
      try { els.moDetailsBody.setAttribute('data-current-json', encodeURIComponent(JSON.stringify(it))); } catch (_) { }
      try { window.__moCurrentItem = it; } catch (_) { }
      try { window.__moCurrentId = it.id; } catch (_) { }
      try {
        const __hist = (it.history || []).slice().sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        if (__hist.length) {
          const rows = __hist.map(h => {
            const who = (h.fullName || '').trim() || (h.email || 'unknown');
            const when = (new Date(h.timestamp)).toLocaleString();
            const act = h.action || 'update';
            return `<div class=\"d-flex align-items-start gap-2 py-1\"><i class=\"fa-regular fa-clock mt-1 text-muted\"></i><div><div><strong>${escapeHtml(act)}</strong> by ${escapeHtml(who)}</div><div class=\"text-muted small\">${escapeHtml(when)}</div></div></div>`;
          }).join('');
          els.moDetailsBody.insertAdjacentHTML('beforeend', `<div class=\"small text-muted mt-3\">Edit History</div><div class=\"border rounded p-2 bg-light small\">${rows}</div>`);
        }
      } catch (_) { }
    }
    moDetailsModal.show();
  }

  // Memo wiring
  if (els.moAddBtn) els.moAddBtn.addEventListener('click', moOnAdd);
  if (els.moForm) els.moForm.addEventListener('submit', moOnSubmit);
  if (els.moSearchInput) els.moSearchInput.addEventListener('input', debounce(() => { moCurrentPage = 1; renderMo(); }, 200));
  if (els.moStatusFilter) els.moStatusFilter.addEventListener('change', () => { moCurrentPage = 1; renderMo(); });
  if (els.moTbody) {
    els.moTbody.addEventListener('click', (ev) => {
      const eBtn = ev.target.closest('.mo-edit'); if (eBtn) { ev.preventDefault(); moOpenEdit(eBtn.getAttribute('data-id')); return; }
      const dBtn = ev.target.closest('.mo-del'); if (dBtn) { ev.preventDefault(); moDelete(dBtn.getAttribute('data-id')); return; }
      const tr = ev.target.closest('tr'); if (!tr) return; const cell = ev.target.closest('td'); if (cell === tr.lastElementChild) return; moOpenDetails(tr.getAttribute('data-id'));
    });
  }
  // Global wrappers
  window.moEdit = (id) => { try { moOpenEdit(id); } catch (_) { } };
  window.moDelete = (id) => { try { moDelete(id); } catch (_) { } };
  window.moOpenDetails = (id) => { try { moOpenDetails(id); } catch (_) { } };

  // ---- Firestore integration for shared data (Office Orders, Memoranda) ----
  try {
    const db = window.firebase?.firestore?.();
    if (db) {
      // Year filter setup - reduces reads from ~thousands to ~hundreds per year
      const yearFilter = document.getElementById('docRegYearFilter');
      let currentDocRegYear = new Date().getFullYear().toString();
      let availableYears = new Set();
      let yearDropdownInitialized = false;

      // Function to update year dropdown with available years from data
      function updateYearDropdown() {
        if (!yearFilter || availableYears.size === 0) return;

        const sortedYears = Array.from(availableYears).sort((a, b) => b - a); // Descending
        const currentSelection = yearFilter.value || currentDocRegYear;

        let optionsHtml = '';
        for (const y of sortedYears) {
          const isSelected = y.toString() === currentSelection;
          optionsHtml += `<option value="${y}"${isSelected ? ' selected' : ''}>${y}</option>`;
        }

        // Only update if we have options
        if (optionsHtml) {
          yearFilter.innerHTML = optionsHtml;
          yearDropdownInitialized = true;
        }
      }

      // Function to extract years from document dates and update dropdown
      function extractYearsFromDocs(docs) {
        let updated = false;
        for (const doc of docs) {
          const docDate = doc.docDate || '';
          if (docDate && docDate.length >= 4) {
            const year = parseInt(docDate.substring(0, 4), 10);
            if (year >= 2000 && year <= 2100 && !availableYears.has(year)) {
              availableYears.add(year);
              updated = true;
            }
          }
        }
        if (updated) updateYearDropdown();
      }

      // Initially set current year as default until data loads
      if (yearFilter) {
        yearFilter.innerHTML = `<option value="${currentDocRegYear}" selected>${currentDocRegYear}</option>`;
      }

      // Unsubscribe functions for all collections
      let unsubNotices = null;
      let unsubOutgoing = null;
      let unsubIncoming = null;
      let unsubMemo = null;
      let unsubOfficeOrder = null;
      let unsubTravelOrder = null;
      let unsubMinutes = null;

      // One-time query to get all years from all collections (for dropdown population)
      async function fetchAvailableYears() {
        try {
          const collections = [
            'docreg_notices', 'docreg_outgoing', 'docreg_incoming',
            'docreg_memoranda', 'docreg_office_orders', 'docreg_travel_orders', 'docreg_minutes'
          ];
          for (const colName of collections) {
            const snap = await db.collection(colName).orderBy('docDate', 'desc').limit(500).get();
            const docs = snap.docs.map(d => d.data());
            extractYearsFromDocs(docs);
          }
        } catch (err) {
          console.warn('Could not fetch available years:', err);
        }
      }
      fetchAvailableYears();

      // Helper function to create year-filtered query
      function getYearQuery(collection) {
        const year = yearFilter?.value || currentDocRegYear;
        const startOfYear = `${year}-01-01`;
        const endOfYear = `${year}-12-31`;
        return collection
          .where('docDate', '>=', startOfYear)
          .where('docDate', '<=', endOfYear)
          .orderBy('docDate', 'desc');
      }

      // Notice of Meeting realtime
      let noItems = [];
      const noCol = db.collection('docreg_notices');
      function subscribeNotices() {
        if (unsubNotices) unsubNotices();
        unsubNotices = getYearQuery(noCol).onSnapshot((snap) => {
          noItems = snap.docs.map(d => ({ id: d.id, ...(d.data() || {}) }));
          try { noLoad = function () { return noItems.slice(); }; } catch (_) { }
          try { renderNo(); } catch (_) { }
        }, err => console.error('Notices subscription error:', err));
      }
      subscribeNotices();
      // Rebind Notice submit to Firestore
      if (els.noForm) {
        try { els.noForm.removeEventListener('submit', noOnSubmit); } catch (_) { }
        els.noForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
          if (!els.noForm.checkValidity()) { els.noForm.reportValidity(); return; }
          const replyTo = els.noForm.getAttribute('data-reply-to') || '';
          const lines = (els.noPartyBlock?.value || '').split(/\r?\n/).map(s => s.trim());
          const sigBlock = (els.noSignatoryBlock?.value || '').split(/\r?\n/).map(s => s.trim());
          const payload = {
            refNo: (els.noRefNo?.value || '').trim(),
            controlNo: (els.noControlNo?.value || '').trim(),
            docDate: els.noDocDate?.value || '',
            meetDate: els.noMeetDate?.value || '',
            meetTime: els.noMeetTime?.value || '',
            name: lines[0] || '',
            agency: lines[1] || '',
            subject: (els.noSubject?.value || '').trim(),
            signatory: (els.noSignatory?.value || '').trim(),
            signatoryDesignation: (sigBlock[0] || (els.noSignatoryDesignation?.value || '').trim()),
            signatoryDepartment: (sigBlock[1] || ''),
            venue: (els.noVenue?.value || '').trim(),
            docLink: (els.noDocLink?.value || '').trim(),
            remarks: (els.noRemarks?.value || '').trim(),
          };
          // Only add createdAt for new documents, not edits
          if (!els.noId?.value) { payload.createdAt = new Date().toISOString(); }
          if (replyTo) { payload.replyToIncomingId = replyTo; }
          const id = els.noId?.value || undefined;
          try {
            const ref = id ? noCol.doc(id) : noCol.doc();
            let __existingHistory = []; try { const snap = await ref.get(); __existingHistory = snap.exists ? ((snap.data() || {}).history || []) : []; } catch (_) { }
            const __user = firebase.auth().currentUser; const __email = __user?.email || 'unknown'; const __full = (__user?.displayName || '').trim(); const __action = id ? 'edit' : 'create';
            const __history = [...__existingHistory, { email: __email, fullName: __full, timestamp: new Date().toISOString(), action: __action }];
            await ref.set({ id: ref.id, history: __history, ...payload }, { merge: true });
            if (els.noId) els.noId.value = '';
            // Firestore-based reply linkage to Incoming
            if (replyTo) {
              try {
                els.noForm.removeAttribute('data-reply-to');
                const incRef = db.collection('docreg_incoming').doc(replyTo);
                const link = { id: ref.id, type: 'notice', date: payload.docDate };
                await incRef.set({ replies: window.firebase.firestore.FieldValue.arrayUnion(link) }, { merge: true });
              } catch (_) { }
            }
            noModal.hide();
          } catch (err) { alert('Failed to save notice: ' + (err?.message || err)); }
        });
      }
      // Rebind Notice delete to Firestore
      try { noDelete = async function (id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } if (!id) return; const ok = confirm('Delete notice? This cannot be undone.'); if (!ok) return; await noCol.doc(id).delete(); }; } catch (_) { }
      // Outgoing Letters realtime
      let outItems = [];
      const outCol = db.collection('docreg_outgoing');
      function subscribeOutgoing() {
        if (unsubOutgoing) unsubOutgoing();
        unsubOutgoing = getYearQuery(outCol).onSnapshot((snap) => {
          outItems = snap.docs.map(d => ({ id: d.id, ...(d.data() || {}) }));
          try { loadItems = function () { return outItems.slice(); }; } catch (_) { }
          try { render(); } catch (_) { }
        }, err => console.error('Outgoing subscription error:', err));
      }
      subscribeOutgoing();
      // Rebind Outgoing submit to Firestore
      if (els.form) {
        try { els.form.removeEventListener('submit', onSubmit); } catch (_) { }
        els.form.addEventListener('submit', async (e) => {
          e.preventDefault();
          if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
          if (!els.form.checkValidity()) { els.form.reportValidity(); return; }
          const replyTo = els.form.getAttribute('data-reply-to') || '';
          const partyLines = (els.olPartyBlock?.value || '').split(/\r?\n/).map(s => s.trim());
          const legacyName = partyLines[0] || '';
          const legacyDesignation = partyLines[1] || '';
          const legacyAgency = partyLines[2] || '';
          const legacyAddress = partyLines[3] || '';
          const toName = (els.olToName?.value || '').trim() || legacyName;
          const toPosition = (els.olToPosition?.value || '').trim() || legacyDesignation;
          const toAgency = (els.olToAgency?.value || '').trim() || legacyAgency;
          const toAddress = (els.olToAddress?.value || '').trim() || legacyAddress;
          const fromName = (els.olFromName?.value || '').trim();
          const fromDesignation = (els.olFromDesignation?.value || '').trim();
          const fromAgency = (els.olFromAgency?.value || '').trim();
          const fromAddress = (els.olFromAddress?.value || '').trim();
          const payload = {
            refNo: (els.refNo?.value || '').trim(),
            controlNo: (els.controlNo?.value || '').trim(),
            docDate: els.docDate?.value || '',
            // legacy display fields mirror TO
            name: toName,
            designation: toPosition,
            agency: toAgency,
            address: toAddress,
            // structured fields
            toName, toPosition, toAgency, toAddress,
            fromName, fromDesignation, fromAgency, fromAddress,
            subject: (els.subject?.value || '').trim(),
            preparedBy: (els.preparedBy?.value || '').trim(),
            docLink: (els.docLink?.value || '').trim(),
            deliveryVia: (els.deliveryVia?.value || '').trim(),
            status: (els.status?.value || ''),
            remarks: (els.remarks?.value || '').trim(),
            signatory: (els.signatory?.value || '').trim(),
          };
          // Add optional fields only if they have values (avoid undefined in Firestore)
          if (replyTo) { payload.replyToIncomingId = replyTo; }
          if (!els.id?.value) { payload.createdAt = new Date().toISOString(); }

          const id = els.id?.value || undefined;
          try {
            const ref = id ? outCol.doc(id) : outCol.doc();
            let __existingHistory = []; try { const snap = await ref.get(); __existingHistory = snap.exists ? ((snap.data() || {}).history || []) : []; } catch (_) { }
            const __user = firebase.auth().currentUser; const __email = __user?.email || 'unknown'; const __full = (__user?.displayName || '').trim(); const __action = id ? 'edit' : 'create';
            const __history = [...__existingHistory, { email: __email, fullName: __full, timestamp: new Date().toISOString(), action: __action }];
            await ref.set({ id: ref.id, history: __history, ...payload }, { merge: true });
            if (els.id) els.id.value = '';
            // Firestore-based reply linkage
            if (replyTo) {
              try {
                els.form.removeAttribute('data-reply-to');
                const incRef = db.collection('docreg_incoming').doc(replyTo);
                const link = { id: ref.id, type: 'outgoing', date: payload.docDate };
                await incRef.set({ replies: window.firebase.firestore.FieldValue.arrayUnion(link) }, { merge: true });
              } catch (_) { }
            }
            olModal.hide();
          } catch (err) { alert('Failed to save outgoing letter: ' + (err?.message || err)); }
        });
      }
      // Rebind Outgoing delete to Firestore (also unlink from Incoming.replies)
      try {
        deleteItem = async function (id, opts = {}) {
          const { skipConfirm = false } = opts || {};
          if (!isAdmin) { alert('Only the Administrator can delete.'); return; }
          if (!id) return;
          const ok = skipConfirm ? true : confirm('Delete document? This cannot be undone.');
          if (!ok) return;
          try {
            const snap = await outCol.doc(id).get();
            const data = snap.exists ? (snap.data() || {}) : {};
            const incId = data.replyToIncomingId || '';
            let affectedIncId = '';
            if (incId) {
              try {
                const incRef = db.collection('docreg_incoming').doc(incId);
                const incSnap = await incRef.get();
                if (incSnap.exists) {
                  const inc = incSnap.data() || {};
                  const curr = Array.isArray(inc.replies) ? inc.replies : [];
                  const next = curr.filter(r => r && r.id !== id);
                  await incRef.set({ replies: next }, { merge: true });
                  affectedIncId = incId;
                }
              } catch (_) { }
            } else {
              // Fallback: scan a limited set of incoming docs to remove stray reply entries by id
              try {
                const qSnap = await db.collection('docreg_incoming').orderBy('docDate', 'desc').limit(500).get();
                for (const d of qSnap.docs) {
                  const inc = d.data() || {}; const arr = Array.isArray(inc.replies) ? inc.replies : [];
                  if (arr.some(r => r && r.id === id)) {
                    const next = arr.filter(r => r && r.id !== id);
                    await d.ref.set({ replies: next }, { merge: true });
                    affectedIncId = d.id; break;
                  }
                }
              } catch (_) { }
            }
            // If the Incoming Details modal is open for this incoming, refresh it so UI reflects removal
            try {
              const openId = els.ilDetailsBody?.getAttribute('data-current-id') || '';
              if (affectedIncId && openId === affectedIncId) { ilOpenDetails(affectedIncId); }
            } catch (_) { }
          } catch (_) { }
          await outCol.doc(id).delete();
        };
      } catch (_) { }

      // Incoming Letters realtime
      let inItems = [];
      const inCol = db.collection('docreg_incoming');
      function subscribeIncoming() {
        if (unsubIncoming) unsubIncoming();
        unsubIncoming = getYearQuery(inCol).onSnapshot((snap) => {
          inItems = snap.docs.map(d => ({ id: d.id, ...(d.data() || {}) }));
          try { ilLoad = function () { return inItems.slice(); }; } catch (_) { }
          try { renderIl(); } catch (_) { }
        }, err => console.error('Incoming subscription error:', err));
      }
      subscribeIncoming();
      // Rebind Incoming submit to Firestore
      if (els.ilForm) {
        try { els.ilForm.removeEventListener('submit', ilOnSubmit); } catch (_) { }
        els.ilForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
          if (!els.ilForm.checkValidity()) { els.ilForm.reportValidity(); return; }
          const toName = (els.ilToName?.value || '').trim();
          const toDesignation = (els.ilToDesignation?.value || '').trim();
          const toAgency = (els.ilToAgency?.value || '').trim();
          const toAddress = (els.ilToAddress?.value || '').trim();
          const fromName = (els.ilFromName?.value || '').trim();
          const fromDesignation = (els.ilFromDesignation?.value || '').trim();
          const fromAgency = (els.ilFromAgency?.value || '').trim();
          const fromAddress = (els.ilFromAddress?.value || '').trim();
          const payload = {
            refNo: (els.ilRefNo?.value || '').trim(),
            docDate: els.ilDocDate?.value || '',
            // legacy display block maps to TO by default
            name: toName,
            designation: toDesignation,
            agency: toAgency,
            address: toAddress,
            // structured fields preserved
            toName, toDesignation, toAgency, toAddress,
            fromName, fromDesignation, fromAgency, fromAddress,
            subject: (els.ilSubject?.value || '').trim(),
            category: (els.ilCategory?.value || '').trim(),
            receivedDate: (els.ilReceivedDate?.value || '').trim(),
            docLink: (els.ilDocLink?.value || '').trim(),
            deliveryVia: (els.ilDeliveryVia?.value || '').trim(),
            assignedTo: (els.ilAssignedTo?.value || '').trim(),
            status: (els.ilStatus?.value || ''),
            remarks: (els.ilRemarks?.value || '').trim(),
          };
          // Only add createdAt for new documents, not edits
          if (!els.ilId?.value) { payload.createdAt = new Date().toISOString(); }
          const id = els.ilId?.value || undefined;
          try {
            const ref = id ? inCol.doc(id) : inCol.doc();
            let __existingHistory = []; try { const snap = await ref.get(); __existingHistory = snap.exists ? ((snap.data() || {}).history || []) : []; } catch (_) { }
            const __user = firebase.auth().currentUser; const __email = __user?.email || 'unknown'; const __full = (__user?.displayName || '').trim(); const __action = id ? 'edit' : 'create';
            const __history = [...__existingHistory, { email: __email, fullName: __full, timestamp: new Date().toISOString(), action: __action }];
            await ref.set({ id: ref.id, history: __history, ...payload }, { merge: true });
            if (els.ilId) els.ilId.value = '';
            ilModal.hide();
          } catch (err) { alert('Failed to save incoming letter: ' + (err?.message || err)); }
        });
      }
      // Rebind Incoming delete to Firestore
      try { ilDelete = async function (id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } if (!id) return; const ok = confirm('Delete document? This cannot be undone.'); if (!ok) return; await inCol.doc(id).delete(); }; } catch (_) { }

      // Rebind reply linkage helpers to work with Firestore
      try {
        finalizeReplyLinkage = async function (createdOutgoing) {
          const incomingId = els.form.getAttribute('data-reply-to');
          if (!incomingId) return;
          try { els.form.removeAttribute('data-reply-to'); } catch (_) { }
          try {
            // ensure outgoing carries replyToIncomingId
            await outCol.doc(createdOutgoing.id).set({ replyToIncomingId: incomingId }, { merge: true });
          } catch (_) { }
          try {
            const incRef = inCol.doc(incomingId);
            const link = { id: createdOutgoing.id, type: 'outgoing', date: createdOutgoing.docDate };
            await incRef.set({ replies: window.firebase.firestore.FieldValue.arrayUnion(link) }, { merge: true });
          } catch (_) { }
        };
      } catch (_) { }

      // Memoranda realtime
      let moItems = [];
      const moCol = db.collection('docreg_memoranda');
      function subscribeMemo() {
        if (unsubMemo) unsubMemo();
        unsubMemo = getYearQuery(moCol).onSnapshot((snap) => {
          moItems = snap.docs.map(d => ({ id: d.id, ...(d.data() || {}) }));
          // Override loader to use snapshot
          try { moLoad = function () { return moItems.slice(); }; } catch (_) { }
          try { renderMo(); } catch (_) { }
        }, err => console.error('Memo subscription error:', err));
      }
      subscribeMemo();
      // Rebind Memorandum submit to Firestore
      if (els.moForm) {
        try { els.moForm.removeEventListener('submit', moOnSubmit); } catch (_) { }
        els.moForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
          if (!els.moForm.checkValidity()) { els.moForm.reportValidity(); return; }
          const party = (els.moPartyBlock?.value || '').split(/\r?\n/).map(s => s.trim());
          const toName = (els.moToName?.value || '').trim() || party[0] || '';
          const toDesignation = (els.moToDesignation?.value || '').trim() || party[1] || '';
          const toDepartment = (els.moToDepartment?.value || '').trim() || party[2] || '';
          const fromName = (els.moFromName?.value || '').trim() || '';
          const fromDesignation = (els.moFromDesignation?.value || '').trim() || '';
          const fromDepartment = (els.moFromDepartment?.value || '').trim() || '';
          const replyTo = els.moForm?.getAttribute('data-reply-to') || '';
          const payload = {
            refNo: (els.moRefNo?.value || '').trim(),
            controlNo: (els.moControlNo?.value || '').trim(),
            docDate: els.moDocDate?.value || '',
            // Legacy display fields mapped to TO
            name: toName,
            designation: toDesignation,
            agency: toDepartment,
            address: party[3] || '',
            // New structured fields
            toName,
            toDesignation,
            toDepartment,
            fromName,
            fromDesignation,
            fromDepartment,
            subject: (els.moSubject?.value || '').trim(),
            preparedBy: (els.moPreparedBy?.value || '').trim(),
            docLink: (els.moDocLink?.value || '').trim(),
            signatory: (els.moSignatory?.value || '').trim(),
            status: els.moStatus?.value || '',
            remarks: (els.moRemarks?.value || '').trim(),
          };
          // Only add createdAt and replyToIncomingId for new documents or when actually set
          if (!els.moId?.value) { payload.createdAt = new Date().toISOString(); }
          if (replyTo) { payload.replyToIncomingId = replyTo; }
          const id = els.moId?.value || undefined;
          try {
            const ref = id ? moCol.doc(id) : moCol.doc();
            await ref.set({ id: ref.id, ...payload }, { merge: true });
            if (els.moId) els.moId.value = '';
            // Firestore-based reply linkage to Incoming
            if (replyTo) {
              try {
                els.moForm.removeAttribute('data-reply-to');
                const incRef = db.collection('docreg_incoming').doc(replyTo);
                const link = { id: ref.id, type: 'memo', date: payload.docDate };
                await incRef.set({ replies: window.firebase.firestore.FieldValue.arrayUnion(link) }, { merge: true });
              } catch (_) { }
            }
            moModal.hide();
          } catch (err) { alert('Failed to save memorandum: ' + (err?.message || err)); }
        });
      }
      // Rebind Memorandum delete to Firestore (also unlink from Incoming.replies)
      try {
        moDelete = async function (id) {
          if (!isAdmin) { alert('Only the Administrator can delete.'); return; }
          if (!id) return;
          const ok = confirm('Delete memorandum? This cannot be undone.');
          if (!ok) return;
          try {
            const snap = await moCol.doc(id).get();
            const data = snap.exists ? (snap.data() || {}) : {};
            const incId = data.replyToIncomingId || '';
            let affectedIncId = '';
            if (incId) {
              try {
                const incRef = db.collection('docreg_incoming').doc(incId);
                const incSnap = await incRef.get();
                if (incSnap.exists) {
                  const inc = incSnap.data() || {};
                  const curr = Array.isArray(inc.replies) ? inc.replies : [];
                  const next = curr.filter(r => r && r.id !== id);
                  await incRef.set({ replies: next }, { merge: true });
                  affectedIncId = incId;
                }
              } catch (_) { }
            } else {
              // Fallback scan if memo lacks linkage
              try {
                const qSnap = await db.collection('docreg_incoming').orderBy('docDate', 'desc').limit(500).get();
                for (const d of qSnap.docs) {
                  const inc = d.data() || {}; const arr = Array.isArray(inc.replies) ? inc.replies : [];
                  if (arr.some(r => r && r.id === id)) {
                    const next = arr.filter(r => r && r.id !== id);
                    await d.ref.set({ replies: next }, { merge: true });
                    affectedIncId = d.id; break;
                  }
                }
              } catch (_) { }
            }
            // Refresh open Incoming details if needed
            try {
              const openId = els.ilDetailsBody?.getAttribute('data-current-id') || '';
              if (affectedIncId && openId === affectedIncId) { ilOpenDetails(affectedIncId); }
            } catch (_) { }
          } catch (_) { }
          await moCol.doc(id).delete();
        };
      } catch (_) { }

      // Office Orders realtime
      let ooItems = [];
      const ooCol = db.collection('docreg_office_orders');
      function subscribeOfficeOrder() {
        if (unsubOfficeOrder) unsubOfficeOrder();
        unsubOfficeOrder = getYearQuery(ooCol).onSnapshot((snap) => {
          ooItems = snap.docs.map(d => ({ id: d.id, ...(d.data() || {}) }));
          try { ooLoad = function () { return ooItems.slice(); }; } catch (_) { }
          try { renderOo(); } catch (_) { }
        }, err => console.error('Office Order subscription error:', err));
      }
      subscribeOfficeOrder();
      // Rebind Office Order submit to Firestore
      if (els.ooForm) {
        try { els.ooForm.removeEventListener('submit', ooOnSubmit); } catch (_) { }
        els.ooForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
          const mode = ooGetMode?.() || 'range';
          // Enforce correct required/disabled flags based on current mode before validation
          ooSetMode?.(mode);
          ooSyncNoEnd?.();
          // Per-mode validation
          if (mode === 'range') {
            const s = normalizeYmd(els.ooIncStart?.value || '');
            const eVal = normalizeYmd(els.ooIncEnd?.value || '');
            if (!s) { els.ooIncStart?.focus(); els.ooForm.reportValidity(); return; }
            if (!(els.ooNoEnd?.checked) && !eVal) { els.ooIncEnd?.focus(); els.ooForm.reportValidity(); return; }
          } else if (mode === 'single') {
            const sd = normalizeYmd(els.ooSingleDate?.value || '');
            if (!sd) { els.ooSingleDate?.focus(); els.ooForm.reportValidity(); return; }
          } else if (mode === 'multi') {
            if (!(ooMultiDates?.length)) { alert('Please add at least one date.'); els.ooMultiDateInput?.focus(); return; }
          }
          if (!els.ooForm.checkValidity()) { els.ooForm.reportValidity(); return; }
          let incStart = '', incEnd = '', noEnd = false, incDates = [], incMode = mode;
          if (mode === 'range') {
            incStart = normalizeYmd(els.ooIncStart?.value || '');
            noEnd = !!els.ooNoEnd?.checked;
            incEnd = noEnd ? '' : normalizeYmd(els.ooIncEnd?.value || '');
            incDates = [];
          } else if (mode === 'single') {
            const sd = normalizeYmd(els.ooSingleDate?.value || '');
            incStart = sd; incEnd = sd; noEnd = false; incDates = [sd];
          } else if (mode === 'multi') {
            incDates = (ooMultiDates || []).slice();
            incStart = ''; incEnd = ''; noEnd = false;
          }
          const payload = {
            refNo: (els.ooRefNo?.value || '').trim(),
            docDate: els.ooDocDate?.value || '',
            subject: (els.ooSubject?.value || '').trim(),
            incMode,
            incStart,
            incEnd,
            noEnd,
            incDates,
            participants: (els.ooParticipants?.value || '').trim(),
            docLink: (els.ooDocLink?.value || '').trim(),
          };
          // Only add createdAt for new documents, not edits
          if (!els.ooId?.value) { payload.createdAt = new Date().toISOString(); }
          const id = els.ooId?.value || undefined;
          try {
            const ref = id ? ooCol.doc(id) : ooCol.doc();
            let __existingHistory = []; try { const snap = await ref.get(); __existingHistory = snap.exists ? ((snap.data() || {}).history || []) : []; } catch (_) { }
            const __user = firebase.auth().currentUser; const __email = __user?.email || 'unknown'; const __full = (__user?.displayName || '').trim(); const __action = id ? 'edit' : 'create';
            const __history = [...__existingHistory, { email: __email, fullName: __full, timestamp: new Date().toISOString(), action: __action }];
            await ref.set({ id: ref.id, history: __history, ...payload }, { merge: true });
            if (els.ooId) els.ooId.value = '';
            ooModal.hide();
          } catch (err) { alert('Failed to save office order: ' + (err?.message || err)); }
        });
      }
      // Rebind Office Order delete to Firestore
      try { ooDelete = async function (id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } if (!id) return; const ok = confirm('Delete office order? This cannot be undone.'); if (!ok) return; await ooCol.doc(id).delete(); }; } catch (_) { }

      // Travel Orders realtime
      let toItems = [];
      const toCol = db.collection('docreg_travel_orders');
      function subscribeTravelOrder() {
        if (unsubTravelOrder) unsubTravelOrder();
        unsubTravelOrder = getYearQuery(toCol).onSnapshot((snap) => {
          toItems = snap.docs.map(d => ({ id: d.id, ...(d.data() || {}) }));
          try { toLoad = function () { return toItems.slice(); }; } catch (_) { }
          try { renderTo(); } catch (_) { }
        }, err => console.error('Travel Order subscription error:', err));
      }
      subscribeTravelOrder();
      // Rebind Travel Order submit to Firestore
      if (els.toForm) {
        try { els.toForm.removeEventListener('submit', toOnSubmit); } catch (_) { }
        els.toForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
          const mode = toGetMode?.() || 'range';
          toSetMode?.(mode);
          toSyncNoEnd?.();
          // per-mode validation
          if (mode === 'range') {
            const s = normalizeYmd(els.toIncStart?.value || '');
            const eVal = normalizeYmd(els.toIncEnd?.value || '');
            if (!s) { els.toIncStart?.focus(); els.toForm.reportValidity(); return; }
            if (!(els.toNoEnd?.checked) && !eVal) { els.toIncEnd?.focus(); els.toForm.reportValidity(); return; }
          } else if (mode === 'single') {
            const sd = normalizeYmd(els.toSingleDate?.value || '');
            if (!sd) { els.toSingleDate?.focus(); els.toForm.reportValidity(); return; }
          } else if (mode === 'multi') {
            if (!(toMultiDates?.length)) { alert('Please add at least one date.'); els.toMultiDateInput?.focus(); return; }
          }
          if (!els.toForm.checkValidity()) { els.toForm.reportValidity(); return; }
          let incStart = '', incEnd = '', noEnd = false, incDates = [], incMode = mode;
          if (mode === 'range') {
            incStart = normalizeYmd(els.toIncStart?.value || '');
            noEnd = !!els.toNoEnd?.checked;
            incEnd = noEnd ? '' : normalizeYmd(els.toIncEnd?.value || '');
          } else if (mode === 'single') {
            const sd = normalizeYmd(els.toSingleDate?.value || '');
            incStart = sd; incEnd = sd; incDates = [sd];
          } else if (mode === 'multi') {
            incDates = (toMultiDates || []).slice();
          }
          const payload = {
            refNo: (els.toRefNo?.value || '').trim(),
            docDate: els.toDocDate?.value || '',
            subject: (els.toSubject?.value || '').trim(),
            incMode,
            incStart,
            incEnd,
            noEnd,
            incDates,
            participants: (els.toParticipants?.value || '').trim(),
            docLink: (els.toDocLink?.value || '').trim(),
          };
          // Only add createdAt for new documents, not edits
          if (!els.toId?.value) { payload.createdAt = new Date().toISOString(); }
          const id = els.toId?.value || undefined;
          try {
            const ref = id ? toCol.doc(id) : toCol.doc();
            let __existingHistory = []; try { const snap = await ref.get(); __existingHistory = snap.exists ? ((snap.data() || {}).history || []) : []; } catch (_) { }
            const __user = firebase.auth().currentUser; const __email = __user?.email || 'unknown'; const __full = (__user?.displayName || '').trim(); const __action = id ? 'edit' : 'create';
            const __history = [...__existingHistory, { email: __email, fullName: __full, timestamp: new Date().toISOString(), action: __action }];
            await ref.set({ id: ref.id, history: __history, ...payload }, { merge: true });
            if (els.toId) els.toId.value = '';
            toModal.hide();
          } catch (err) { alert('Failed to save travel order: ' + (err?.message || err)); }
        });
      }
      // Rebind Travel Order delete to Firestore
      try { toDelete = async function (id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } if (!id) return; const ok = confirm('Delete travel order? This cannot be undone.'); if (!ok) return; await toCol.doc(id).delete(); }; } catch (_) { }

      // Minutes of Meeting realtime
      let miItems = [];
      const miCol = db.collection('docreg_minutes');
      function subscribeMinutes() {
        if (unsubMinutes) unsubMinutes();
        unsubMinutes = getYearQuery(miCol).onSnapshot((snap) => {
          miItems = snap.docs.map(d => ({ id: d.id, ...(d.data() || {}) }));
          try { miLoad = function () { return miItems.slice(); }; } catch (_) { }
          try { renderMi(); } catch (_) { }
        }, err => console.error('Minutes subscription error:', err));
      }
      subscribeMinutes();
      // Rebind Minutes submit to Firestore
      if (els.miForm) {
        try { els.miForm.removeEventListener('submit', miOnSubmit); } catch (_) { }
        els.miForm.addEventListener('submit', async (e) => {
          e.preventDefault();
          if (!canEditDocReg) { alert('View-only access. Contact your administrator.'); return; }
          if (!els.miForm.checkValidity()) { els.miForm.reportValidity(); return; }
          const payload = {
            controlNo: (els.miControlNo?.value || '').trim(),
            docDate: els.miDocDate?.value || '',
            meetDate: els.miMeetDate?.value || '',
            subject: (els.miSubject?.value || '').trim(),
            signatory: (els.miSignatory?.value || '').trim(),
            agenda: (els.miAgenda?.value || '').trim(),
            preparedBy: (els.miPreparedBy?.value || '').trim(),
            docLink: (els.miDocLink?.value || '').trim(),
            remarks: (els.miRemarks?.value || '').trim(),
          };
          // Only add createdAt for new documents, not edits
          if (!els.miId?.value) { payload.createdAt = new Date().toISOString(); }
          const id = els.miId?.value || undefined;
          try {
            const ref = id ? miCol.doc(id) : miCol.doc();
            let __existingHistory = []; try { const snap = await ref.get(); __existingHistory = snap.exists ? ((snap.data() || {}).history || []) : []; } catch (_) { }
            const __user = firebase.auth().currentUser; const __email = __user?.email || 'unknown'; const __full = (__user?.displayName || '').trim(); const __action = id ? 'edit' : 'create';
            const __history = [...__existingHistory, { email: __email, fullName: __full, timestamp: new Date().toISOString(), action: __action }];
            await ref.set({ id: ref.id, history: __history, ...payload }, { merge: true });
            if (els.miId) els.miId.value = '';
            miModal.hide();
          } catch (err) { alert('Failed to save minutes: ' + (err?.message || err)); }
        });
      }
      // Rebind Minutes delete to Firestore
      try { miDelete = async function (id) { if (!isAdmin) { alert('Only the Administrator can delete.'); return; } if (!id) return; const ok = confirm('Delete minutes? This cannot be undone.'); if (!ok) return; await miCol.doc(id).delete(); }; } catch (_) { }

      // Year filter change handler - re-subscribe all collections when year changes
      if (yearFilter) {
        yearFilter.addEventListener('change', () => {
          console.log('Document Registry: Switching to year', yearFilter.value);
          // Re-subscribe all 7 collections with new year filter
          subscribeNotices();
          subscribeOutgoing();
          subscribeIncoming();
          subscribeMemo();
          subscribeOfficeOrder();
          subscribeTravelOrder();
          subscribeMinutes();
        });
      }
    }
  } catch (_) { }

  // Global wrappers for inline incoming action buttons
  window.ilEdit = (id) => { try { ilOpenEdit(id); } catch (_) { } };
  window.ilDelete = (id) => { try { ilDelete(id); } catch (_) { } };
  window.ilOpenDetails = (id) => { try { ilOpenDetails(id); } catch (_) { } };

  // Default to Incoming section on load
  switchSection('incoming');

  // Wire up
  if (els.addBtn) els.addBtn.addEventListener('click', onAddClick);
  if (els.form) els.form.addEventListener('submit', onSubmit);
  // Filters wiring
  if (els.searchInput) els.searchInput.addEventListener('input', debounce(() => { currentPage = 1; render(); }, 200));
  if (els.statusFilter) els.statusFilter.addEventListener('change', () => { currentPage = 1; render(); });
  if (els.deliveryFilter) els.deliveryFilter.addEventListener('change', () => { currentPage = 1; render(); });

  // Row actions and row click to open details
  if (els.tbody) {
    els.tbody.addEventListener('click', (ev) => {
      const editBtn = ev.target.closest('.ol-edit');
      if (editBtn) { ev.preventDefault(); openEdit(editBtn.getAttribute('data-id')); return; }
      const delBtn = ev.target.closest('.ol-del');
      if (delBtn) { ev.preventDefault(); window.olDelete && window.olDelete(delBtn.getAttribute('data-id')); return; }
      const tr = ev.target.closest('tr');
      if (!tr) return;
      // Ignore clicks on action cell
      const cell = ev.target.closest('td');
      const actionCell = tr.lastElementChild;
      if (cell === actionCell) return;
      openDetails(tr.getAttribute('data-id'));
    });
  }

  // Initial render
  render();
  try { renderMo(); } catch (_) { }
  try { renderNo(); } catch (_) { }
})();

// ========== Word Document Generation Functions (using Docxtemplater) ==========
(function () {
  'use strict';

  // Templates for different document types
  const TEMPLATES = {
    outgoing: 'assets/outgoing_letter_template.docx',
    memorandum: 'assets/memorandum_template.docx',
    notice: 'assets/notice_of_meeting_template.docx',
    project: 'assets/site_inspection_template.docx',
    // Add more templates as needed
  };

  // Dynamically load a script if not already loaded
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
      const s = document.createElement('script');
      s.src = src;
      s.async = true;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('Failed to load script ' + src));
      document.head.appendChild(s);
    });
  }

  // Ensure PizZip and Docxtemplater are loaded
  async function ensureDocxDeps() {
    const DocxTemplater = window.Docxtemplater || window.docxtemplater;
    if (window.PizZip && DocxTemplater) return;

    if (!window.PizZip) {
      try { await loadScript('assets/vendor/pizzip.min.js'); } catch (_) { }
      if (!window.PizZip) {
        try { await loadScript('https://cdn.jsdelivr.net/npm/pizzip@3.1.7/dist/pizzip.min.js'); } catch (_) { }
        if (!window.PizZip) {
          await loadScript('https://unpkg.com/pizzip@3.1.7/dist/pizzip.min.js');
        }
      }
    }
    if (!(window.Docxtemplater || window.docxtemplater)) {
      try { await loadScript('assets/vendor/docxtemplater.js'); } catch (_) { }
      if (!(window.Docxtemplater || window.docxtemplater)) {
        try { await loadScript('https://cdn.jsdelivr.net/npm/docxtemplater@3.22.2/build/docxtemplater.js'); } catch (_) { }
        if (!(window.Docxtemplater || window.docxtemplater)) {
          await loadScript('https://unpkg.com/docxtemplater@3.22.2/build/docxtemplater.js');
        }
      }
    }
  }

  // Helper to format date nicely as '10 December 2025'
  function formatDocDate(dateStr) {
    if (!dateStr) return 'n/a';
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return dateStr;

      const day = d.getDate();
      const month = d.toLocaleDateString('en-US', { month: 'long' });
      const year = d.getFullYear();
      return `${day} ${month} ${year}`;
    } catch (_) { return dateStr; }
  }

  // Return 'n/a' for empty values
  function valOrNA(v) {
    if (v === null || v === undefined) return 'n/a';
    if (typeof v === 'string') {
      const t = v.trim();
      return t ? v : 'n/a';
    }
    return v;
  }

  // Load template as ArrayBuffer
  async function loadTemplateArrayBuffer(templatePath) {
    const res = await fetch(templatePath, { cache: 'no-store' });
    if (!res.ok) {
      throw new Error('Failed to load template: ' + templatePath + ' (' + res.status + ')');
    }
    return await res.arrayBuffer();
  }

  // Generate document from template using docxtemplater
  async function generateFromTemplate(templatePath, data) {
    await ensureDocxDeps();

    const DocxTemplater = window.Docxtemplater || window.docxtemplater;
    if (!(window.PizZip && DocxTemplater)) {
      throw new Error('Docxtemplater or PizZip not loaded');
    }

    const content = await loadTemplateArrayBuffer(templatePath);
    const zip = new window.PizZip(content);
    const doc = new DocxTemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      nullGetter(part) {
        console.warn('Missing template var:', part.value);
        return 'n/a';
      }
    });

    if (typeof doc.compile === 'function') {
      doc.compile();
    }

    try {
      if (typeof doc.resolveData === 'function') {
        await doc.resolveData(data);
      } else {
        doc.setData(data);
      }
      doc.render();
    } catch (err) {
      console.error('Docxtemplater Render Error:', err);
      if (err.properties && err.properties.errors) {
        err.properties.errors.forEach(e => console.error('Docxtemplater detailed error:', e));
      }
      throw new Error(err.message || 'Template render failed');
    }

    return doc.getZip().generate({ type: 'blob' });
  }

  // Download blob as file
  function downloadBlob(blob, filename) {
    if (window.saveAs) {
      window.saveAs(blob, filename);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 100);
    }
  }

  // Outgoing Letter Word generation
  window.olGenerateFromTemplate = async function (item) {
    if (!item) { alert('No document data available.'); return; }

    try {
      console.log('Generating Outgoing Letter from template...', item);
      const formattedDate = formatDocDate(item.docDate);
      console.log('Document Date Formatted:', formattedDate);

      // Map item fields to template placeholders
      const data = {
        RefNo: valOrNA(item.refNo),
        ControlNo: valOrNA(item.controlNo),
        DocDate: formattedDate,
        Date: formattedDate,
        DocumentDate: formattedDate,
        DocumentDateLong: formattedDate, // Exact variable used in template

        PreparedBy: valOrNA(item.preparedBy),
        Name: valOrNA(item.name),
        Designation: valOrNA(item.designation),
        Agency: valOrNA(item.agency),
        Address: valOrNA(item.address),
        Subject: valOrNA(item.subject),
        Signatory: valOrNA(item.signatory),
        DeliveryVia: valOrNA(item.deliveryVia),
        Status: valOrNA(item.status),
        Remarks: valOrNA(item.remarks),

        // Common alternate names and verbose fields
        ReferenceNo: valOrNA(item.refNo),
        RecipientName: valOrNA(item.name),
        RecipientDesignation: valOrNA(item.designation),
        RecipientAgency: valOrNA(item.agency),
        RecipientAddress: valOrNA(item.address),
      };

      const blob = await generateFromTemplate(TEMPLATES.outgoing, data);
      const filename = `Outgoing_${item.refNo || item.controlNo || 'Letter'}_${item.docDate || 'undated'}.docx`;
      downloadBlob(blob, filename.replace(/[\\/:*?"<>|]/g, '_'));
      console.log('Outgoing Letter downloaded successfully');
    } catch (err) {
      console.error('Error generating document:', err);
      alert('Failed to generate Word document: ' + (err.message || err));
    }
  };

  // Memorandum Word generation
  window.moGenerateFromTemplate = async function (item) {
    if (!item) { alert('No document data available.'); return; }

    try {
      const formattedDate = formatDocDate(item.docDate);
      const data = {
        RefNo: valOrNA(item.refNo),
        ControlNo: valOrNA(item.controlNo),
        DocDate: formattedDate,
        Date: formattedDate,
        DocumentDate: formattedDate,
        DocumentDateLong: formattedDate,

        // TO fields (Mixed Case)
        ToName: valOrNA(item.toName || item.name),
        ToDesignation: valOrNA(item.toDesignation || item.designation),
        ToDepartment: valOrNA(item.toDepartment || item.agency || item.department),
        ToAgency: valOrNA(item.toDepartment || item.agency),
        ToAddress: valOrNA(item.toAddress || item.address),

        // TO fields (Template Specific Uppercase Prefix)
        TOName: valOrNA(item.toName || item.name),
        TODesignation: valOrNA(item.toDesignation || item.designation),
        TODepartment: valOrNA(item.toDepartment || item.agency || item.department),

        // FROM fields (Mixed Case)
        FromName: valOrNA(item.fromName),
        FromDesignation: valOrNA(item.fromDesignation),
        FromDepartment: valOrNA(item.fromDepartment || item.fromAgency || item.fromOffice),
        FromAgency: valOrNA(item.fromDepartment || item.fromAgency),

        // FROM fields (Template Specific Uppercase Prefix)
        FROMName: valOrNA(item.fromName),
        FROMDesignation: valOrNA(item.fromDesignation),
        FROMDepartment: valOrNA(item.fromDepartment || item.fromAgency || item.fromOffice),

        Subject: valOrNA(item.subject),
        Status: valOrNA(item.status),
        Remarks: valOrNA(item.remarks),
      };

      const blob = await generateFromTemplate(TEMPLATES.memorandum, data);
      const filename = `Memo_${item.refNo || 'Memorandum'}_${item.docDate || 'undated'}.docx`;
      downloadBlob(blob, filename.replace(/[\\/:*?"<>|]/g, '_'));
    } catch (err) {
      console.error('Error generating document:', err);
      alert('Failed to generate Word document: ' + (err.message || err));
    }
  };

  // Notice of Meeting Word generation
  window.noGenerateFromTemplate = async function (item) {
    if (!item) { alert('No document data available.'); return; }

    try {
      const formattedDate = formatDocDate(item.docDate);
      const meetDate = formatDocDate(item.meetDate);

      const data = {
        RefNo: valOrNA(item.refNo),
        ControlNo: valOrNA(item.controlNo),
        DocDate: formattedDate,
        Date: formattedDate,
        DocumentDate: formattedDate,
        DocumentDateLong: formattedDate,

        MeetDate: meetDate,
        MeetTime: valOrNA(item.meetTime),
        TimeofMeeting: valOrNA(item.meetTime), // Exact var from log

        Venue: valOrNA(item.venue),
        Subject: valOrNA(item.subject),

        Name: valOrNA(item.name),
        Agency: valOrNA(item.agency),

        Signatory: valOrNA(item.signatory),
        SIGNATORY: valOrNA(item.signatory), // Exact var from log

        // Contextual Designation/Department
        Designation: valOrNA(item.designation || item.signatoryDesignation),
        Department: valOrNA(item.department || item.signatoryDepartment || item.agency),

        Status: valOrNA(item.status),
        Remarks: valOrNA(item.remarks),
      };

      const blob = await generateFromTemplate(TEMPLATES.notice, data);
      const filename = `Notice_${item.refNo || item.controlNo || 'Meeting'}_${item.meetDate || item.docDate || 'undated'}.docx`;
      downloadBlob(blob, filename.replace(/[\\/:*?"<>|]/g, '_'));
    } catch (err) {
      console.error('Error generating document:', err);
      alert('Failed to generate Word document: ' + (err.message || err));
    }
  };

  // Fallback implementations for other types (uses simple HTML-to-Word)
  function generateSimpleWordDoc(title, fields, filename) {
    const rows = fields.map(f => `<tr><th width="30%">${f.label}</th><td>${f.value}</td></tr>`).join('');
    const html = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head><meta charset="UTF-8">
      <style>@page{size:letter;margin:1in}body{font-family:Calibri,Arial,sans-serif;font-size:11pt}h1{font-size:14pt;text-align:center}table{width:100%;border-collapse:collapse;margin:12pt 0}td,th{border:1px solid #999;padding:6pt 8pt}th{background:#e0e0e0;font-weight:bold}</style>
      </head><body><h1>${title}</h1><table>${rows}</table></body></html>`;
    const blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
    downloadBlob(blob, filename.replace(/\.docx$/, '.doc'));
  }

  // Incoming Letter
  window.ilGenerateFromTemplate = function (item) {
    if (!item) { alert('No document data available.'); return; }
    const fields = [
      { label: 'Reference No.', value: valOrNA(item.refNo) },
      { label: 'Document Date', value: formatDocDate(item.docDate) },
      { label: 'Date Received', value: formatDocDate(item.receivedDate) },
      { label: 'From', value: valOrNA(item.name) },
      { label: 'Designation', value: valOrNA(item.designation) },
      { label: 'Agency', value: valOrNA(item.agency) },
      { label: 'Subject', value: valOrNA(item.subject) },
      { label: 'Status', value: valOrNA(item.status) },
      { label: 'Remarks', value: valOrNA(item.remarks) },
    ];
    generateSimpleWordDoc('INCOMING LETTER', fields, `Incoming_${item.refNo || 'Letter'}_${item.docDate || 'undated'}.doc`);
  };

  // Office Order
  window.ooGenerateFromTemplate = function (item) {
    if (!item) { alert('No document data available.'); return; }
    const fields = [
      { label: 'Order No.', value: valOrNA(item.orderNo || item.refNo) },
      { label: 'Document Date', value: formatDocDate(item.docDate) },
      { label: 'Subject', value: valOrNA(item.subject) },
      { label: 'Signatory', value: valOrNA(item.signatory) },
      { label: 'Status', value: valOrNA(item.status) },
      { label: 'Remarks', value: valOrNA(item.remarks) },
    ];
    generateSimpleWordDoc('OFFICE ORDER', fields, `OfficeOrder_${item.orderNo || item.refNo || 'Order'}_${item.docDate || 'undated'}.doc`);
  };

  // Travel Order
  window.toGenerateFromTemplate = function (item) {
    if (!item) { alert('No document data available.'); return; }
    const fields = [
      { label: 'Order No.', value: valOrNA(item.orderNo || item.refNo) },
      { label: 'Document Date', value: formatDocDate(item.docDate) },
      { label: 'Traveler Name', value: valOrNA(item.name) },
      { label: 'Destination', value: valOrNA(item.destination) },
      { label: 'Purpose', value: valOrNA(item.purpose || item.subject) },
      { label: 'Travel Date', value: formatDocDate(item.travelDate) },
      { label: 'Return Date', value: formatDocDate(item.returnDate) },
      { label: 'Status', value: valOrNA(item.status) },
      { label: 'Remarks', value: valOrNA(item.remarks) },
    ];
    generateSimpleWordDoc('TRAVEL ORDER', fields, `TravelOrder_${item.orderNo || item.refNo || 'Order'}_${item.docDate || 'undated'}.doc`);
  };

  // Minutes of Meeting
  window.miGenerateFromTemplate = function (item) {
    if (!item) { alert('No document data available.'); return; }
    const fields = [
      { label: 'Reference No.', value: valOrNA(item.refNo) },
      { label: 'Meeting Date', value: formatDocDate(item.meetDate || item.docDate) },
      { label: 'Meeting Time', value: valOrNA(item.meetTime) },
      { label: 'Venue', value: valOrNA(item.venue) },
      { label: 'Subject/Agenda', value: valOrNA(item.subject) },
      { label: 'Attendees', value: valOrNA(item.attendees) },
      { label: 'Status', value: valOrNA(item.status) },
      { label: 'Remarks', value: valOrNA(item.remarks) },
    ];
    generateSimpleWordDoc('MINUTES OF MEETING', fields, `Minutes_${item.refNo || 'Meeting'}_${item.meetDate || item.docDate || 'undated'}.doc`);
  };

  // ---- UI Connectors ----

  // Memorandum - using stored global item
  window.moPrintCurrent = async function () {
    let item = window.__moCurrentItem;
    if (!item) {
      try {
        const json = document.getElementById('moDetailsBody')?.getAttribute('data-current-json');
        if (json) item = JSON.parse(decodeURIComponent(json));
      } catch (_) { }
    }
    await window.moGenerateFromTemplate(item);
  };

  // Notice - capture ID and fetch
  let currentNoticeId = null;
  const origNoOpenDetails = window.noOpenDetails;
  window.noOpenDetails = function (id) {
    currentNoticeId = id;
    if (origNoOpenDetails) origNoOpenDetails(id);
  };

  window.noPrintCurrent = async function () {
    if (!currentNoticeId) { alert('No document selected.'); return; }
    try {
      // Direct Firestore fetch as we cannot access internal cache
      const db = window.firebase?.firestore?.();
      if (!db) { alert('Database not connected'); return; }
      const snap = await db.collection('docreg_notices').doc(currentNoticeId).get();
      if (!snap.exists) { alert('Document no longer exists.'); return; }
      const item = { id: snap.id, ...(snap.data() || {}) };
      await window.noGenerateFromTemplate(item);
    } catch (err) {
      console.error(err);
      alert('Failed to generate notice document: ' + err.message);
    }
  };
  console.log('[docreg] Word generation functions registered');
})();
