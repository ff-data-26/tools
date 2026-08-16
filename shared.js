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

  return { parseSpreadsheet, setupDropZone, val, makeKey, escHtml, showError };
})();
