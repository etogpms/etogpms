// Shared utility helpers for dashboard pages.
// Exposed on window.AppUtils for reuse across modules and inline scripts.
(function (window) {
  if (window.AppUtils) {
    return;
  }

  function fmtNum(val) {
    const num = Number(val);
    return Number.isNaN(num) ? (val || '') : num.toLocaleString();
  }

  function escapeHtml(s) {
    const str = String(s == null ? '' : s);
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function escapeAttr(s) {
    return escapeHtml(s);
  }

  function normalizeYmd(v) {
    if (!v) return '';
    try {
      const s = String(v).trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
      const d = new Date(s);
      if (Number.isNaN(d.getTime())) return '';
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const da = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${da}`;
    } catch (_) {
      return '';
    }
  }

  function formatDateUI(v) {
    if (!v) return '';
    const str = String(v).trim();
    const parts = str.split('-');
    if (parts.length === 3 && parts[0].length === 4) {
      const m = parseInt(parts[1], 10);
      const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      if (m >= 1 && m <= 12) {
        return `${parts[0]}-${months[m-1]}-${parts[2]}`;
      }
    }
    return str;
  }

  function safeDateCompare(a, b) {
    const aa = normalizeYmd(a) || '';
    const bb = normalizeYmd(b) || '';
    if (aa === bb) return 0;
    return aa < bb ? -1 : 1;
  }

  function canonText(s) {
    return (s || '').toString().trim().replace(/\s+/g, ' ').toUpperCase();
  }

  function todayYmd() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const da = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${da}`;
  }

  function jan1OfYear(y) {
    return `${y}-01-01`;
  }

  function dec31OfYear(y) {
    return `${y}-12-31`;
  }

  function debounce(fn, wait) {
    let t;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), wait);
    };
  }

  function canonBnaq(s) {
    return (s || '').toString().trim().toUpperCase().replace(/\s*-\s*/g, '-').replace(/\s+/g, ' ');
  }

  function daysInMonthStr(ym) {
    if (!ym || ym.length < 7) return 30;
    const y = parseInt(ym.slice(0, 4), 10);
    const m = parseInt(ym.slice(5, 7), 10);
    return new Date(y, m, 0).getDate();
  }

  function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error('Failed to read blob'));
      reader.onloadend = () => {
        const base64String = reader.result.split(',')[1];
        resolve(base64String);
      };
      reader.readAsDataURL(blob);
    });
  }

  async function downloadFileMobile(data, filename, isWorkbook = false) {
    try {
      const cap = window.Capacitor;
      if (!cap || !cap.Plugins) {
        throw new Error('Capacitor plugins not available');
      }
      const { Filesystem, Share } = cap.Plugins;
      if (!Filesystem || !Share) {
        throw new Error('Capacitor Filesystem or Share plugin not loaded');
      }

      let base64Data = '';
      if (isWorkbook) {
        if (!window.XLSX) throw new Error('SheetJS library not loaded');
        base64Data = window.XLSX.write(data, { bookType: 'xlsx', type: 'base64' });
      } else if (data instanceof Blob) {
        base64Data = await blobToBase64(data);
      } else if (typeof data === 'string') {
        const blob = new Blob([data], { type: 'text/plain;charset=utf-8' });
        base64Data = await blobToBase64(blob);
      } else {
        throw new Error('Unsupported data format for mobile download');
      }

      const writeResult = await Filesystem.writeFile({
        path: filename,
        data: base64Data,
        directory: 'CACHE'
      });

      await Share.share({
        title: filename,
        url: writeResult.uri,
        dialogTitle: 'Save / Share File'
      });
    } catch (err) {
      console.error('Mobile download failed:', err);
      alert('Failed to save file: ' + (err.message || String(err)));
    }
  }

  function setupMobileOverrides() {
    const isNative = window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform();
    if (!isNative) return;

    let originalSaveAs = window.saveAs;
    if (originalSaveAs) {
      window.saveAs = function (blob, filename) {
        downloadFileMobile(blob, filename, false);
      };
    } else {
      let currentSaveAs = undefined;
      Object.defineProperty(window, 'saveAs', {
        get() { return currentSaveAs; },
        set(fn) {
          currentSaveAs = function (blob, filename) {
            downloadFileMobile(blob, filename, false);
          };
        },
        configurable: true
      });
    }

    let originalXLSX = window.XLSX;
    if (originalXLSX && typeof originalXLSX.writeFile === 'function') {
      originalXLSX.writeFile = function (wb, filename) {
        downloadFileMobile(wb, filename, true);
      };
    } else {
      let currentXLSX = originalXLSX;
      Object.defineProperty(window, 'XLSX', {
        get() { return currentXLSX; },
        set(obj) {
          currentXLSX = obj;
          if (currentXLSX && typeof currentXLSX.writeFile === 'function') {
            currentXLSX.writeFile = function (wb, filename) {
              downloadFileMobile(wb, filename, true);
            };
          }
        },
        configurable: true
      });
    }
  }

  setupMobileOverrides();

  const utils = {
    fmtNum,
    escapeHtml,
    escapeAttr,
    normalizeYmd,
    safeDateCompare,
    canonText,
    todayYmd,
    jan1OfYear,
    dec31OfYear,
    debounce,
    canonBnaq,
    daysInMonthStr,
    formatDateUI,
    downloadFileMobile,
  };

  window.AppUtils = Object.freeze(utils);
})(window);
