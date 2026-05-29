// Page and modal scroll helpers: shows floating up/down buttons for long content and scrollable modals.
// Exposes UIShell.initScrollControls().
(function (window) {
  if (window.UIShell?.initScrollControls) return;

  function initScrollControls(opts = {}) {
    const scrollTopBtn = document.getElementById(opts.pageTopId || 'scrollTopBtn');
    const scrollBottomBtn = document.getElementById(opts.pageBottomId || 'scrollBottomBtn');

    function toggleScrollButtons() {
      const scrolled = document.documentElement.scrollTop || document.body.scrollTop;
      const maxScroll = document.documentElement.scrollHeight - window.innerHeight;
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
      scrollTopBtn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
      scrollBottomBtn.addEventListener('click', () => window.scrollTo({ top: document.documentElement.scrollHeight, behavior: 'smooth' }));
    }

    // Modal scroll arrows
    const modalUp = document.getElementById(opts.modalUpId || 'modalScrollUp');
    const modalDown = document.getElementById(opts.modalDownId || 'modalScrollDown');
    const suppressedIds = opts.suppressedModalIds || [
      'detailsModal',
      'deepwellModal',
      'reforestationModal',
      'reforestationDetailsModal',
      'usersModal',
      'serviceUpdateModal',
      'bnaqEntriesModal',
      'bnaqModal',
      'bnaqGroupEditModal',
    ];
    let activeModalBody = null;

    const getActiveModalBody = () => {
      const shown = document.querySelector('.modal.show');
      if (!shown) return null;
      return shown.querySelector('.modal-dialog-scrollable .modal-body') || shown.querySelector('.modal-body');
    };

    const positionModalScrollBtns = () => {
      if (!modalUp || !modalDown) return;
      const dlg = document.querySelector('.modal.show .modal-dialog');
      if (!dlg) return;
      const rect = dlg.getBoundingClientRect();
      const left = Math.min(rect.right - 44, window.innerWidth - 52);
      modalUp.style.left = left + 'px';
      modalDown.style.left = left + 'px';
      modalUp.style.right = 'auto';
      modalDown.style.right = 'auto';
    };

    const updateModalScrollBtns = () => {
      if (!modalUp || !modalDown) return;
      const shown = document.querySelector('.modal.show');
      const suppressed = shown && suppressedIds.includes(shown.id);
      if (suppressed) {
        modalUp.style.display = 'none';
        modalDown.style.display = 'none';
        if (scrollTopBtn) scrollTopBtn.style.display = 'none';
        if (scrollBottomBtn) scrollBottomBtn.style.display = 'none';
        activeModalBody = null;
        return;
      }
      activeModalBody = getActiveModalBody();
      if (!activeModalBody) {
        modalUp.style.display = 'none';
        modalDown.style.display = 'none';
        return;
      }
      positionModalScrollBtns();
      const scrolled = activeModalBody.scrollTop;
      const max = activeModalBody.scrollHeight - activeModalBody.clientHeight;
      modalUp.style.display = scrolled > 50 ? 'flex' : 'none';
      modalDown.style.display = scrolled < max - 50 ? 'flex' : 'none';
    };

    if (modalUp) {
      modalUp.addEventListener('click', () => {
        const body = getActiveModalBody();
        if (body) body.scrollBy({ top: -300, behavior: 'smooth' });
      });
    }
    if (modalDown) {
      modalDown.addEventListener('click', () => {
        const body = getActiveModalBody();
        if (body) body.scrollBy({ top: 300, behavior: 'smooth' });
      });
    }

    window.addEventListener('resize', updateModalScrollBtns);
    document.addEventListener('shown.bs.modal', updateModalScrollBtns);
    document.addEventListener('hidden.bs.modal', () => {
      if (!modalUp || !modalDown) return;
      modalUp.style.display = 'none';
      modalDown.style.display = 'none';
      activeModalBody = null;
      try { toggleScrollButtons(); } catch (_) { /* noop */ }
    });
    document.addEventListener('scroll', () => {
      if (document.querySelector('.modal.show')) updateModalScrollBtns();
    }, true);
  }

  window.UIShell = Object.freeze({
    ...(window.UIShell || {}),
    initScrollControls,
  });
})(window);
