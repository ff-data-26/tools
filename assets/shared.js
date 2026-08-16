/* ============================================================
   tools/assets/shared.js
   Shared helpers for the tools hub. Requires SheetJS (xlsx) to
   already be loaded on the page for parseSpreadsheet to work.
   ============================================================ */

const TOOLS = (function () {

  // Parse an uploaded .xlsx / .xls / .csv into { rows, headers }.
  // cb(err, rows, headers)
  function parseSpreadsheet(file, cb) {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        const headers = rows.length ? Object.keys(rows[0]) : [];
        cb(null, rows, headers);
      } catch (err) {
        cb('Could not parse file: ' + err.message);
      }
    };
    reader.readAsBinaryString(file);
  }

  // Wire drag/drop + click-to-upload behavior onto a drop-zone element.
  // zoneEl: the .drop-zone container. onFile(file) fires on either path.
  function setupDropZone(zoneEl, onFile) {
    const input = zoneEl.querySelector('input[type=file]');
    if (input) {
      input.addEventListener('change', ev => {
        const file = ev.target.files[0];
        if (file) onFile(file);
      });
    }
    zoneEl.addEventListener('dragover', ev => {
      ev.preventDefault();
      zoneEl.classList.add('drag-over');
    });
    zoneEl.addEventListener('dragleave', () => zoneEl.classList.remove('drag-over'));
    zoneEl.addEventListener('drop', ev => {
      ev.preventDefault();
      zoneEl.classList.remove('drag-over');
      const file = ev.dataTransfer.files[0];
      if (file) onFile(file);
    });
  }

  // Normalize a cell value to a trimmed string (dates -> ISO date).
  function val(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).trim();
  }

  // Build a composite key by joining selected column values with a
  // null-byte separator (won't collide with real data).
  function makeKey(row, cols) {
    return cols.map(c => val(row[c])).join('\u0000');
  }

  function escHtml(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // Show/hide the standard error banner. Pass '' to hide.
  function showError(bannerId, msg) {
    const b = document.getElementById(bannerId);
    if (!b) return;
    b.textContent = msg;
    b.style.display = msg ? 'block' : 'none';
  }

  // ── Theme switcher ──────────────────────────────────────────
  // Injects a small fixed-position swatch picker. Selection is
  // stored in localStorage under 'tools-theme' and read on every
  // page load (see the inline snippet in each page's <head>,
  // which applies it before first paint to avoid a flash).
  const THEMES = [
    { id: 'default', label: 'Default', swatch: '#eae7df' },
    { id: 'sky',     label: 'Sky',     swatch: '#cfe2ef' },
    { id: 'sage',    label: 'Sage',    swatch: '#d6e4cd' },
    { id: 'dusk',    label: 'Dusk',    swatch: '#ddd2ee' }
  ];

  function applyTheme(id, persist) {
    if (id === 'default') document.documentElement.removeAttribute('data-theme');
    else document.documentElement.setAttribute('data-theme', id);
    if (persist) localStorage.setItem('tools-theme', id);
    const bar = document.querySelector('.theme-switcher');
    if (bar) {
      Array.from(bar.children).forEach((el, i) => el.classList.toggle('active', THEMES[i].id === id));
    }
  }

  function initThemeSwitcher() {
    const current = document.documentElement.getAttribute('data-theme') || 'default';
    const bar = document.createElement('div');
    bar.className = 'theme-switcher';
    THEMES.forEach(t => {
      const btn = document.createElement('button');
      btn.className = 'theme-swatch' + (t.id === current ? ' active' : '');
      btn.style.background = t.swatch;
      btn.title = t.label;
      btn.setAttribute('aria-label', 'Theme: ' + t.label);
      btn.addEventListener('click', () => applyTheme(t.id, true));
      bar.appendChild(btn);
    });
    document.body.appendChild(bar);
  }

  return { parseSpreadsheet, setupDropZone, val, makeKey, escHtml, showError, initThemeSwitcher };
})();
