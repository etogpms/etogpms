// Sidebar toggle/collapse behavior for dashboard layout.
// Exposes UIShell.initSidebarToggle to wire DOM events and restore persisted state.
(function (window) {
  if (window.UIShell?.initSidebarToggle) return;

  function initSidebarToggle(opts = {}) {
    const appWrapper = document.getElementById(opts.appWrapperId || 'appWrapper');
    const sidebarToggle = document.getElementById(opts.toggleId || 'sidebarToggle');
    const sidebarEl = document.getElementById(opts.sidebarId || 'sidebar');
    const sidebarOverlay = document.getElementById(opts.overlayId || 'sidebarOverlay');
    if (!appWrapper || !sidebarToggle || !sidebarEl) return;

    const KEY = 'ui.sidebar.collapsed';

    const setIcon = (collapsed) => {
      const i = sidebarToggle.querySelector('i');
      if (!i) return;
      i.className = collapsed ? 'fa-solid fa-chevron-right' : 'fa-solid fa-chevron-left';
      sidebarToggle.setAttribute('aria-label', collapsed ? 'Show sidebar' : 'Hide sidebar');
      sidebarToggle.setAttribute('title', collapsed ? 'Show sidebar' : 'Hide sidebar');
    };

    const applyState = (collapsed) => {
      const isMobile = window.matchMedia('(max-width: 991.98px)').matches;
      if (isMobile) {
        if (collapsed) {
          sidebarEl.classList.remove('sidebar-mobile-visible');
          if (sidebarOverlay) sidebarOverlay.classList.remove('show');
        } else {
          sidebarEl.classList.add('sidebar-mobile-visible');
          if (sidebarOverlay) sidebarOverlay.classList.add('show');
        }
      } else {
        if (collapsed) {
          appWrapper.classList.add('sidebar-collapsed');
          sidebarEl.classList.add('sidebar-hidden');
        } else {
          appWrapper.classList.remove('sidebar-collapsed');
          sidebarEl.classList.remove('sidebar-hidden');
        }
        sidebarEl.classList.remove('sidebar-mobile-visible');
        if (sidebarOverlay) sidebarOverlay.classList.remove('show');
      }
      setIcon(collapsed);
    };

    const initial = localStorage.getItem(KEY) === '1';
    const isMobileInit = window.matchMedia('(max-width: 991.98px)').matches;
    applyState(isMobileInit ? true : initial);

    sidebarToggle.addEventListener('click', () => {
      const isMobile = window.matchMedia('(max-width: 991.98px)').matches;
      let collapsed;
      if (isMobile) {
        collapsed = sidebarEl.classList.contains('sidebar-mobile-visible');
      } else {
        collapsed = !appWrapper.classList.contains('sidebar-collapsed');
      }
      applyState(collapsed);
      if (!isMobile) localStorage.setItem(KEY, collapsed ? '1' : '0');
    });

    if (sidebarOverlay) {
      sidebarOverlay.addEventListener('click', () => {
        applyState(true);
      });
    }

    // Auto-close on nav link click (mobile only)
    sidebarEl.querySelectorAll('.nav-link, #documentRegistryBtn').forEach(link => {
      link.addEventListener('click', () => {
        if (window.matchMedia('(max-width: 991.98px)').matches) {
          applyState(true);
        }
      });
    });

    window.addEventListener('resize', () => {
      const isMobile = window.matchMedia('(max-width: 991.98px)').matches;
      if (!isMobile) {
        const stored = localStorage.getItem(KEY) === '1';
        applyState(stored);
      } else {
        if (!sidebarEl.classList.contains('sidebar-mobile-visible')) {
          if (sidebarOverlay) sidebarOverlay.classList.remove('show');
        }
      }
    });
  }

  window.UIShell = Object.freeze({
    ...(window.UIShell || {}),
    initSidebarToggle,
  });
})(window);
