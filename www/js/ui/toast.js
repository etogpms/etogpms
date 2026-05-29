// Toast helpers for lightweight, non-blocking notifications.
// Provides UIShell.ensureToastContainer() to create/reuse a Bootstrap toast container.
(function (window) {
  if (window.UIShell?.ensureToastContainer) return;

  function ensureToastContainer() {
    let container = document.querySelector('.toast-container');
    if (container) return container;
    container = document.createElement('div');
    container.className = 'toast-container position-fixed bottom-0 end-0 p-3';
    container.style.zIndex = '1080';
    document.body.appendChild(container);
    return container;
  }

  function showToast(message, type = 'info') {
    const container = ensureToastContainer();
    const id = 'toast-' + Date.now();
    let bgClass = 'text-bg-primary';
    if (type === 'success') bgClass = 'text-bg-success';
    if (type === 'danger' || type === 'error') bgClass = 'text-bg-danger';
    if (type === 'warning') bgClass = 'text-bg-warning text-dark';
    if (type === 'info') bgClass = 'text-bg-info text-dark';

    const div = document.createElement('div');
    div.className = `toast align-items-center ${bgClass} border-0`;
    div.setAttribute('role', 'alert');
    div.setAttribute('aria-live', 'assertive');
    div.setAttribute('aria-atomic', 'true');
    div.id = id;

    div.innerHTML = `
      <div class="d-flex">
        <div class="toast-body">
          ${message}
        </div>
        <button type="button" class="btn-close ${type === 'warning' || type === 'info' ? '' : 'btn-close-white'} me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
      </div>
    `;

    container.appendChild(div);

    if (window.bootstrap?.Toast) {
      const toast = new window.bootstrap.Toast(div, { delay: 3000 });
      toast.show();
      div.addEventListener('hidden.bs.toast', () => div.remove());
    } else {
      // Fallback
      div.style.display = 'block';
      setTimeout(() => div.remove(), 3000);
    }
  }

  window.UIShell = Object.freeze({
    ...(window.UIShell || {}),
    ensureToastContainer,
    showToast,
  });
  window.showToast = showToast;
})(window);
