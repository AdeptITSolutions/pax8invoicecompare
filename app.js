/* ═══════════════════════════════════════════════════
   Pax8 Invoice Comparison Engine
   Grouped by Company → SKU
   ═══════════════════════════════════════════════════ */

(() => {
  'use strict';

  // ── Element refs ──
  const dropzoneA = document.getElementById('dropzone-a');
  const dropzoneB = document.getElementById('dropzone-b');
  const fileInputA = document.getElementById('file-input-a');
  const fileInputB = document.getElementById('file-input-b');
  const fileInfoA = document.getElementById('file-info-a');
  const fileInfoB = document.getElementById('file-info-b');
  const cardA = document.getElementById('upload-card-a');
  const cardB = document.getElementById('upload-card-b');
  const compareBtn = document.getElementById('compare-btn');
  const uploadSection = document.getElementById('upload-section');
  const resultsSection = document.getElementById('results-section');
  const backBtn = document.getElementById('back-btn');
  const searchInput = document.getElementById('search-input');
  const diffThead = document.getElementById('diff-thead');
  const diffTbody = document.getElementById('diff-tbody');
  const keyColSelect = document.getElementById('key-column-select');
  const recompareBtn = document.getElementById('recompare-btn');
  const exportBtn = document.getElementById('export-btn');

  // ── State ──
  let dataA = null;
  let dataB = null;
  let diffResult = null;
  let currentFilter = 'all';

  // ── Columns we care about for comparison ──
  const DISPLAY_COLS = [
    'type', 'sku', 'description', 'quantity',
    'price', 'subtotal', 'total'
  ];

  // Numeric columns (for smart comparison)
  const NUMERIC_COLS = new Set(['quantity', 'price', 'subtotal', 'cost', 'cost_total', 'total', 'sales_tax', 'partner_subtotal', 'amount_due']);

  // ═════════════════════════════════
  //  CSV Parser
  // ═════════════════════════════════
  function parseCSV(text) {
    const lines = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (i + 1 < text.length && text[i + 1] === '"') {
            current += '"'; i++;
          } else { inQuotes = false; }
        } else { current += ch; }
      } else {
        if (ch === '"') { inQuotes = true; }
        else if (ch === '\n' || ch === '\r') {
          if (ch === '\r' && i + 1 < text.length && text[i + 1] === '\n') i++;
          lines.push(current); current = '';
        } else { current += ch; }
      }
    }
    if (current.length > 0) lines.push(current);

    const result = lines.map(line => splitCSVLine(line));
    return result.filter(row => row.some(cell => cell.trim() !== ''));
  }

  function splitCSVLine(line) {
    const fields = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQuotes) {
        if (ch === '"') {
          if (i + 1 < line.length && line[i + 1] === '"') {
            current += '"'; i++;
          } else { inQuotes = false; }
        } else { current += ch; }
      } else {
        if (ch === '"') { inQuotes = true; }
        else if (ch === ',') { fields.push(current.trim()); current = ''; }
        else { current += ch; }
      }
    }
    fields.push(current.trim());
    return fields;
  }

  // ═════════════════════════════════
  //  File handling
  // ═════════════════════════════════
  function handleFile(file, side) {
    if (!file || !file.name.toLowerCase().endsWith('.csv')) {
      alert('Please upload a CSV file.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const text = e.target.result;
      const rows = parseCSV(text);
      if (rows.length < 2) {
        alert('CSV must have at least a header row and one data row.');
        return;
      }

      const headers = rows[0].map(h => h.trim().toLowerCase());
      const dataRows = rows.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = (i < row.length ? row[i] : '').trim();
        });
        return obj;
      });

      const parsed = { headers, rows: dataRows, fileName: file.name };

      if (side === 'a') {
        dataA = parsed;
        showFileInfo(fileInfoA, parsed);
        cardA.classList.add('has-file');
      } else {
        dataB = parsed;
        showFileInfo(fileInfoB, parsed);
        cardB.classList.add('has-file');
      }

      compareBtn.disabled = !(dataA && dataB);
    };
    reader.readAsText(file);
  }

  function showFileInfo(el, parsed) {
    // Count unique companies
    const companyCol = findCol(parsed.headers, ['company_name']);
    const companyCount = companyCol ? new Set(parsed.rows.map(r => (r[companyCol] || '').trim())).size : '?';
    el.innerHTML = `
      <span class="file-name">📄 ${escapeHTML(parsed.fileName)}</span>
      <span class="file-rows">&middot; ${parsed.rows.length} line items &middot; ${companyCount} companies</span>
    `;
  }

  function findCol(headers, candidates) {
    for (const c of candidates) {
      if (headers.includes(c)) return c;
    }
    return null;
  }

  // ═════════════════════════════════
  //  Drag & Drop + Click
  // ═════════════════════════════════
  function setupDropzone(zone, input, side) {
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
    });
    input.addEventListener('change', () => {
      if (input.files.length) handleFile(input.files[0], side);
    });
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', (e) => {
      e.preventDefault(); zone.classList.remove('dragover');
      if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0], side);
    });
  }

  setupDropzone(dropzoneA, fileInputA, 'a');
  setupDropzone(dropzoneB, fileInputB, 'b');

  // ═════════════════════════════════
  //  Comparison Engine — Grouped by Company → SKU
  // ═════════════════════════════════

  function normaliseCompany(name) {
    return (name || '').trim().toUpperCase();
  }

  function groupByCompany(rows) {
    const map = new Map();
    for (const row of rows) {
      const company = normaliseCompany(row.company_name);
      if (!map.has(company)) map.set(company, []);
      map.get(company).push(row);
    }
    return map;
  }

  /**
   * Within a company, aggregate rows by SKU.
   * For each SKU we sum quantity, subtotal, total and keep track of individual line items.
   */
  function aggregateBySku(rows) {
    const map = new Map();
    for (const row of rows) {
      const sku = (row.sku || 'NO-SKU').trim();
      if (!map.has(sku)) {
        map.set(sku, {
          sku,
          description: row.description || '',
          type: row.type || '',
          quantity: 0,
          subtotal: 0,
          total: 0,
          price: row.price || '',
          lines: []
        });
      }
      const agg = map.get(sku);
      agg.quantity += parseFloat(row.quantity) || 0;
      agg.subtotal += parseFloat(row.subtotal) || 0;
      agg.total += parseFloat(row.total) || 0;
      agg.lines.push(row);

      // Keep the subscription description if available (prefer it over prorate)
      if (row.type === 'subscription' && row.description) {
        agg.description = row.description;
        agg.type = row.type;
      }
      if (!agg.description && row.description) {
        agg.description = row.description;
      }
    }
    return map;
  }

  function computeDiff() {
    const companiesA = groupByCompany(dataA.rows);
    const companiesB = groupByCompany(dataB.rows);

    const allCompanyNames = new Set([...companiesA.keys(), ...companiesB.keys()]);
    const sortedCompanies = [...allCompanyNames].sort();

    const companyDiffs = [];

    let totalUnchanged = 0, totalModified = 0, totalAdded = 0, totalRemoved = 0;

    for (const companyName of sortedCompanies) {
      const rowsA = companiesA.get(companyName) || [];
      const rowsB = companiesB.get(companyName) || [];

      const skuMapA = aggregateBySku(rowsA);
      const skuMapB = aggregateBySku(rowsB);

      const allSkus = new Set([...skuMapA.keys(), ...skuMapB.keys()]);
      const skuDiffs = [];

      for (const sku of allSkus) {
        const a = skuMapA.get(sku);
        const b = skuMapB.get(sku);

        if (!a) {
          // Added in B
          skuDiffs.push({ status: 'added', sku, a: null, b, changes: [] });
          totalAdded++;
        } else if (!b) {
          // Removed from A
          skuDiffs.push({ status: 'removed', sku, a, b: null, changes: [] });
          totalRemoved++;
        } else {
          // Both exist — compare aggregated values
          const changes = [];
          const qtyA = roundNum(a.quantity);
          const qtyB = roundNum(b.quantity);
          if (qtyA !== qtyB) changes.push({ field: 'quantity', a: qtyA, b: qtyB });

          const subA = roundNum(a.subtotal);
          const subB = roundNum(b.subtotal);
          if (subA !== subB) changes.push({ field: 'subtotal', a: subA, b: subB });

          const totA = roundNum(a.total);
          const totB = roundNum(b.total);
          if (totA !== totB) changes.push({ field: 'total', a: totA, b: totB });

          // Check if description changed meaningfully (ignore minor wording)
          const descA = a.description.replace(/\[options:.*?\]/g, '').trim();
          const descB = b.description.replace(/\[options:.*?\]/g, '').trim();
          if (descA !== descB) {
            changes.push({ field: 'description', a: a.description, b: b.description });
          }

          if (changes.length > 0) {
            skuDiffs.push({ status: 'modified', sku, a, b, changes });
            totalModified++;
          } else {
            skuDiffs.push({ status: 'unchanged', sku, a, b, changes: [] });
            totalUnchanged++;
          }
        }
      }

      // Display company name from original data (not normalised)
      const displayName = (rowsB.length > 0 ? rowsB[0].company_name : rowsA[0].company_name || '').trim();

      // Determine company-level status
      const hasChanges = skuDiffs.some(s => s.status !== 'unchanged');

      companyDiffs.push({
        companyName: displayName,
        companyNameNorm: companyName,
        skuDiffs,
        hasChanges,
        totalA: rowsA.reduce((s, r) => s + (parseFloat(r.total) || 0), 0),
        totalB: rowsB.reduce((s, r) => s + (parseFloat(r.total) || 0), 0)
      });
    }

    return {
      companies: companyDiffs,
      counts: {
        total: totalUnchanged + totalModified + totalAdded + totalRemoved,
        unchanged: totalUnchanged,
        modified: totalModified,
        added: totalAdded,
        removed: totalRemoved,
        companies: companyDiffs.length,
        companiesChanged: companyDiffs.filter(c => c.hasChanges).length
      }
    };
  }

  function roundNum(n) {
    return Math.round(n * 100) / 100;
  }

  // ═════════════════════════════════
  //  Rendering
  // ═════════════════════════════════
  function renderResults() {
    // Hide key column bar (we use company+sku now)
    document.getElementById('key-column-bar').style.display = 'none';

    updateSummary();
    renderTable();
  }

  function updateSummary() {
    const c = diffResult.counts;
    animateCounter('summary-total', c.total);
    animateCounter('summary-unchanged', c.unchanged);
    animateCounter('summary-modified', c.modified);
    animateCounter('summary-added', c.added);
    animateCounter('summary-removed', c.removed);

    // Update labels
    document.querySelector('#summary-total .summary-label').textContent = `SKU Lines (${c.companies} companies)`;
  }

  function animateCounter(id, target) {
    const el = document.getElementById(id).querySelector('.summary-value');
    const duration = 600;
    const start = performance.now();
    const from = parseInt(el.textContent) || 0;
    function tick(now) {
      const t = Math.min((now - start) / duration, 1);
      const ease = 1 - Math.pow(1 - t, 3);
      el.textContent = Math.round(from + (target - from) * ease);
      if (t < 1) requestAnimationFrame(tick);
    }
    requestAnimationFrame(tick);
  }

  function renderTable() {
    const query = searchInput.value.toLowerCase();

    // Clear
    diffThead.innerHTML = '';
    diffTbody.innerHTML = '';

    // Header row
    const headRow = document.createElement('tr');
    headRow.innerHTML = `
      <th class="row-status"></th>
      <th>Status</th>
      <th>SKU</th>
      <th>Description</th>
      <th style="text-align:right">Qty (A)</th>
      <th style="text-align:right">Qty (B)</th>
      <th style="text-align:right">Total (A)</th>
      <th style="text-align:right">Total (B)</th>
      <th style="text-align:right">Difference</th>
    `;
    diffThead.appendChild(headRow);

    const fragment = document.createDocumentFragment();
    let visibleCount = 0;

    for (const company of diffResult.companies) {
      // Filter at company level
      const matchingSkus = company.skuDiffs.filter(sd => {
        // Status filter
        if (currentFilter !== 'all' && sd.status !== currentFilter) return false;
        // Search filter
        if (query) {
          const haystack = `${company.companyName} ${sd.sku} ${sd.a?.description || ''} ${sd.b?.description || ''}`.toLowerCase();
          if (!haystack.includes(query)) return false;
        }
        return true;
      });

      if (matchingSkus.length === 0) continue;

      // Company header row
      const companyRow = document.createElement('tr');
      companyRow.className = 'company-header-row';
      const diff = roundNum(company.totalB - company.totalA);
      const diffClass = diff > 0 ? 'diff-positive' : diff < 0 ? 'diff-negative' : '';
      const diffSign = diff > 0 ? '+' : '';

      companyRow.innerHTML = `
        <td colspan="6" class="company-name-cell">
          <strong>${escapeHTML(company.companyName || '(No Company)')}</strong>
          <span class="company-sku-count">${matchingSkus.length} SKU${matchingSkus.length !== 1 ? 's' : ''}</span>
        </td>
        <td style="text-align:right; font-weight:600; font-size:.78rem; color:var(--text-muted);">${formatCurrency(company.totalA)}</td>
        <td style="text-align:right; font-weight:600; font-size:.78rem; color:var(--text-muted);">${formatCurrency(company.totalB)}</td>
        <td style="text-align:right; font-weight:700; font-size:.78rem;" class="${diffClass}">${diffSign}${formatCurrency(diff)}</td>
      `;
      fragment.appendChild(companyRow);

      // SKU rows under this company
      for (const sd of matchingSkus) {
        const tr = document.createElement('tr');
        tr.classList.add(`row-${sd.status}`);

        const desc = sd.b?.description || sd.a?.description || '';
        const shortDesc = desc.length > 80 ? desc.substring(0, 77) + '…' : desc;

        const qtyA = sd.a ? formatNum(sd.a.quantity) : '—';
        const qtyB = sd.b ? formatNum(sd.b.quantity) : '—';
        const totA = sd.a ? formatCurrency(sd.a.total) : '—';
        const totB = sd.b ? formatCurrency(sd.b.total) : '—';

        const skuDiff = roundNum((sd.b?.total || 0) - (sd.a?.total || 0));
        const skuDiffClass = skuDiff > 0 ? 'diff-positive' : skuDiff < 0 ? 'diff-negative' : '';
        const skuDiffSign = skuDiff > 0 ? '+' : '';

        // Highlight changed cells
        const qtyChanged = sd.changes.some(c => c.field === 'quantity');
        const totChanged = sd.changes.some(c => c.field === 'total' || c.field === 'subtotal');

        tr.innerHTML = `
          <td class="row-status"><span class="status-dot ${sd.status}"></span></td>
          <td class="status-label status-${sd.status}">${capitalise(sd.status)}</td>
          <td class="sku-cell">${escapeHTML(sd.sku)}</td>
          <td class="desc-cell" title="${escapeHTML(desc)}">${escapeHTML(shortDesc)}</td>
          <td style="text-align:right" class="${qtyChanged ? 'cell-changed-val' : ''}">${qtyA}</td>
          <td style="text-align:right" class="${qtyChanged ? 'cell-changed-val' : ''}">${qtyB}</td>
          <td style="text-align:right" class="${totChanged ? 'cell-changed-val' : ''}">${totA}</td>
          <td style="text-align:right" class="${totChanged ? 'cell-changed-val' : ''}">${totB}</td>
          <td style="text-align:right; font-weight:600;" class="${skuDiffClass}">${skuDiff !== 0 ? skuDiffSign + formatCurrency(skuDiff) : '—'}</td>
        `;

        // Tooltip with change details
        if (sd.changes.length > 0) {
          tr.title = sd.changes.map(c => `${c.field}: ${c.a} → ${c.b}`).join('\n');
        }

        fragment.appendChild(tr);
        visibleCount++;
      }
    }

    diffTbody.appendChild(fragment);

    // Empty state
    if (visibleCount === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = 9;
      td.style.textAlign = 'center';
      td.style.padding = '40px 16px';
      td.style.color = 'var(--text-muted)';
      td.textContent = 'No rows match the current filter / search.';
      tr.appendChild(td);
      diffTbody.appendChild(tr);
    }
  }

  // ═════════════════════════════════
  //  Events
  // ═════════════════════════════════
  compareBtn.addEventListener('click', () => {
    if (!dataA || !dataB) return;
    diffResult = computeDiff();
    uploadSection.classList.add('hidden');
    resultsSection.classList.remove('hidden');
    renderResults();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  backBtn.addEventListener('click', () => {
    resultsSection.classList.add('hidden');
    uploadSection.classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  recompareBtn.addEventListener('click', () => {
    if (!dataA || !dataB) return;
    diffResult = computeDiff();
    renderResults();
  });

  // Filters
  document.querySelectorAll('.filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentFilter = btn.dataset.filter;
      renderTable();
    });
  });

  // Search
  let searchTimeout;
  searchInput.addEventListener('input', () => {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => renderTable(), 200);
  });

  // ═════════════════════════════════
  //  Excel Export
  // ═════════════════════════════════
  exportBtn.addEventListener('click', exportToExcel);

  function exportToExcel() {
    if (!diffResult) return;

    const wb = XLSX.utils.book_new();

    // ── Sheet 1: Summary ──
    const summaryData = [
      ['Pax8 Invoice Comparison — Summary'],
      [],
      ['Metric', 'Count'],
      ['Total SKU Lines', diffResult.counts.total],
      ['Unchanged', diffResult.counts.unchanged],
      ['Modified', diffResult.counts.modified],
      ['Added in Invoice B', diffResult.counts.added],
      ['Removed from Invoice A', diffResult.counts.removed],
      [],
      ['Companies', diffResult.counts.companies],
      ['Companies with Changes', diffResult.counts.companiesChanged],
      [],
      ['File A', dataA.fileName],
      ['File B', dataB.fileName],
      ['Generated', new Date().toLocaleString()]
    ];

    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);

    // Style summary column widths
    wsSummary['!cols'] = [{ wch: 30 }, { wch: 20 }];

    // Merge title row
    wsSummary['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];

    XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

    // ── Sheet 2: Details (grouped by company) ──
    const detailHeaders = [
      'Company', 'Status', 'SKU', 'Description',
      'Qty (A)', 'Qty (B)', 'Qty Change',
      'Total (A)', 'Total (B)', 'Total Change',
      'Changes'
    ];
    const detailRows = [detailHeaders];

    for (const company of diffResult.companies) {
      // Company summary row
      const compDiff = roundNum(company.totalB - company.totalA);
      detailRows.push([
        company.companyName || '(No Company)',
        company.hasChanges ? 'HAS CHANGES' : 'NO CHANGES',
        '', '',
        '', '', '',
        company.totalA,
        company.totalB,
        compDiff,
        ''
      ]);

      // SKU rows
      for (const sd of company.skuDiffs) {
        const qtyA = sd.a ? roundNum(sd.a.quantity) : '';
        const qtyB = sd.b ? roundNum(sd.b.quantity) : '';
        const qtyChg = (sd.a && sd.b) ? roundNum((sd.b?.quantity || 0) - (sd.a?.quantity || 0)) : '';
        const totA = sd.a ? roundNum(sd.a.total) : '';
        const totB = sd.b ? roundNum(sd.b.total) : '';
        const totChg = roundNum((sd.b?.total || 0) - (sd.a?.total || 0));
        const desc = sd.b?.description || sd.a?.description || '';
        const changes = sd.changes.map(c => `${c.field}: ${c.a} → ${c.b}`).join('; ');

        detailRows.push([
          '', // company col blank for SKU rows
          capitalise(sd.status),
          sd.sku,
          desc,
          qtyA, qtyB, qtyChg,
          totA, totB, totChg,
          changes
        ]);
      }

      // Blank spacer row between companies
      detailRows.push([]);
    }

    const wsDetails = XLSX.utils.aoa_to_sheet(detailRows);

    // Column widths
    wsDetails['!cols'] = [
      { wch: 40 }, // Company
      { wch: 12 }, // Status
      { wch: 22 }, // SKU
      { wch: 55 }, // Description
      { wch: 10 }, // Qty A
      { wch: 10 }, // Qty B
      { wch: 12 }, // Qty Change
      { wch: 14 }, // Total A
      { wch: 14 }, // Total B
      { wch: 14 }, // Total Change
      { wch: 40 }, // Changes
    ];

    // Apply number formatting to currency columns
    const currencyCols = [7, 8, 9]; // Total A, Total B, Total Change (0-indexed)
    const numericCols = [4, 5, 6];  // Qty A, Qty B, Qty Change
    for (let r = 1; r < detailRows.length; r++) {
      for (const c of currencyCols) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        if (wsDetails[cellRef] && typeof wsDetails[cellRef].v === 'number') {
          wsDetails[cellRef].z = '$#,##0.00';
        }
      }
      for (const c of numericCols) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        if (wsDetails[cellRef] && typeof wsDetails[cellRef].v === 'number') {
          wsDetails[cellRef].z = '#,##0.####';
        }
      }
    }

    XLSX.utils.book_append_sheet(wb, wsDetails, 'Details by Company');

    // Generate and download
    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Pax8_Invoice_Comparison_${today}.xlsx`);
  }

  // ═════════════════════════════════
  //  Utils
  // ═════════════════════════════════
  function escapeHTML(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function capitalise(s) {
    return s.charAt(0).toUpperCase() + s.slice(1);
  }

  function formatNum(n) {
    const num = typeof n === 'number' ? n : parseFloat(n);
    if (isNaN(num)) return String(n);
    // Show integers without decimal, others with up to 4
    if (Number.isInteger(num)) return num.toLocaleString();
    return num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 4 });
  }

  function formatCurrency(n) {
    const num = typeof n === 'number' ? n : parseFloat(n);
    if (isNaN(num)) return String(n);
    return '$' + num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

})();
