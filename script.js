// ── DATA ─────────────────────────────────────────────
let ftthData = null, dslData = null, irData = null;
let finalData = [], filteredData = [], headers = [];
let sortCol = null, sortDir = 1;
let activeFilter = 'all', searchVal = '';

const PREVIEW = 500;

// ── FILE LOADING ──────────────────────────────────────
function setupDrop(dzId, inputId, type) {
    const dz = document.getElementById(dzId);
    const inp = document.getElementById(inputId);

    inp.addEventListener('change', e => {
        if (e.target.files[0]) loadFile(e.target.files[0], type);
    });
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => {
        e.preventDefault();
        dz.classList.remove('drag-over');
        const f = e.dataTransfer.files[0];
        if (f) { inp.files = e.dataTransfer.files; loadFile(f, type); }
    });
}
setupDrop('dz-ftth', 'file-ftth', 'ftth');
setupDrop('dz-dsl',  'file-dsl',  'dsl');
setupDrop('dz-ir',   'file-ir',   'ir');

function loadFile(file, type) {
    showOverlay(`Loading ${type.toUpperCase()}…`);
    const reader = new FileReader();
    reader.onload = e => {
        const wb   = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });

        const badge = document.getElementById(`${type}-status`);
        badge.innerHTML = `<span class="dot"></span> ${json.length.toLocaleString()} rows`;
        badge.classList.add('ok');
        document.getElementById(`dz-${type}`).classList.add('loaded');

        if (type === 'ftth') ftthData = json;
        if (type === 'dsl')  dslData  = json;
        if (type === 'ir')   irData   = json;
        hideOverlay();
    };
    reader.readAsArrayBuffer(file);
}

// ── CLEAN ─────────────────────────────────────────────
function clean(val) {
    if (!val) return '';
    return String(val).toUpperCase().replace(/FBB/g, '').replace(/FV/g, '').trim();
}

// ── RUN ───────────────────────────────────────────────
document.getElementById('btn-run').addEventListener('click', () => {
    if (!ftthData || !dslData || !irData) {
        alert('Please upload all 3 sheets first.');
        return;
    }
    showOverlay('Running cross-match…');

    setTimeout(() => {
        const dslMap  = new Map();
        const ftthMap = new Map();
        dslData.forEach(r  => { const n = clean(r['Service Number']); if (n) dslMap.set(n, r); });
        ftthData.forEach(r => { const n = clean(r['Service Number']); if (n) ftthMap.set(n, r); });

        finalData = irData.map(irRow => {
            const row = { ...irRow };
            const num = clean(irRow['dsl_number']);
            const cat = String(irRow['case_sub_type'] || '').trim().toLowerCase();

            // DSL match
            const dsl = dslMap.get(num);
            if (dsl) {
                const s = String(dsl['Service Name'] || '').trim();
                row['Current Activity DSL'] = s;
                row['DSL Match'] = (cat === s.toLowerCase()) ? 'The same' : 'Changed';
            } else {
                row['Current Activity DSL'] = 'Not avail';
                row['DSL Match'] = 'Not avail';
            }

            // FTTH match
            const ftth = ftthMap.get(num);
            if (ftth) {
                const s = String(ftth['Service Name'] || '').trim();
                row['Current Activity FTTH'] = s;
                row['FTTH Match'] = (cat === s.toLowerCase()) ? 'The same' : 'Changed';
            } else {
                row['Current Activity FTTH'] = 'Not avail';
                row['FTTH Match'] = 'Not avail';
            }

            return row;
        });

        headers = finalData.length ? Object.keys(finalData[0]) : [];
        updateStats();
        applyFilter();

        document.getElementById('btn-export').style.display = 'inline-block';
        document.getElementById('stats-grid').style.display = 'grid';
        document.getElementById('filter-bar').style.display = 'flex';
        document.getElementById('empty-state').style.display = 'none';
        document.getElementById('table-inner').style.display = 'block';
        document.getElementById('tbl-footer').style.display  = 'flex';
        hideOverlay();
    }, 50);
});

// ── STATS ─────────────────────────────────────────────
function updateStats() {
    const tot  = finalData.length;
    const same = finalData.filter(r => r['FTTH Match'] === 'The same').length;
    const chg  = finalData.filter(r => r['FTTH Match'] === 'Changed').length;
    const na   = finalData.filter(r => r['FTTH Match'] === 'Not avail').length;
    const pct  = n => tot ? Math.round(n / tot * 100) + '%' : '—';

    document.getElementById('s-total').textContent  = tot.toLocaleString();
    document.getElementById('s-same').textContent   = same.toLocaleString();
    document.getElementById('s-chg').textContent    = chg.toLocaleString();
    document.getElementById('s-na').textContent     = na.toLocaleString();
    document.getElementById('s-same-p').textContent = pct(same) + ' of total';
    document.getElementById('s-chg-p').textContent  = pct(chg)  + ' of total';
    document.getElementById('s-na-p').textContent   = pct(na)   + ' of total';
}

// ── FILTER & SEARCH ───────────────────────────────────
document.querySelectorAll('.chip').forEach(c => {
    c.addEventListener('click', () => {
        document.querySelectorAll('.chip').forEach(x => x.classList.remove('active'));
        c.classList.add('active');
        activeFilter = c.dataset.f;
        applyFilter();
    });
});

document.getElementById('search-input').addEventListener('input', e => {
    searchVal = e.target.value.toLowerCase();
    applyFilter();
});

function applyFilter() {
    filteredData = finalData.filter(row => {
        if (activeFilter === 'same' && row['FTTH Match'] !== 'The same')   return false;
        if (activeFilter === 'chg'  && row['FTTH Match'] !== 'Changed')    return false;
        if (activeFilter === 'na'   && row['FTTH Match'] !== 'Not avail')  return false;
        if (searchVal) {
            return Object.values(row).some(v => String(v).toLowerCase().includes(searchVal));
        }
        return true;
    });
    if (sortCol !== null) doSort(false);
    renderTable();
}

// ── SORT ──────────────────────────────────────────────
function doSort(toggle = true) {
    if (toggle) sortDir *= -1;
    const key = headers[sortCol];
    filteredData.sort((a, b) => String(a[key] || '').localeCompare(String(b[key] || '')) * sortDir);
}

// ── RENDER ────────────────────────────────────────────
function renderTable() {
    const thead = document.getElementById('tbl-head');
    const tbody = document.getElementById('tbl-body');

    // Header row
    thead.innerHTML = '<tr>' + headers.map((h, i) => {
        const cls = sortCol === i ? (sortDir === 1 ? 'sort-asc' : 'sort-desc') : '';
        return `<th class="${cls}" data-i="${i}">${h} <span class="sort-arrow"></span></th>`;
    }).join('') + '</tr>';

    thead.querySelectorAll('th').forEach(th => {
        th.addEventListener('click', () => { sortCol = +th.dataset.i; doSort(true); renderTable(); });
    });

    // Data rows (preview only)
    const slice = filteredData.slice(0, PREVIEW);
    const frag  = document.createDocumentFragment();

    slice.forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(h => {
            const td = document.createElement('td');
            const v  = row[h] !== undefined ? String(row[h]) : '';
            if (h === 'FTTH Match' || h === 'DSL Match') {
                if      (v === 'The same')  td.innerHTML = `<span class="badge-same">✔ The same</span>`;
                else if (v === 'Changed')   td.innerHTML = `<span class="badge-chg">⚡ Changed</span>`;
                else if (v === 'Not avail') td.innerHTML = `<span class="badge-na">✕ Not avail</span>`;
                else td.textContent = v;
            } else {
                td.textContent = v;
            }
            tr.appendChild(td);
        });
        frag.appendChild(tr);
    });
    tbody.innerHTML = '';
    tbody.appendChild(frag);

    // Footer
    document.getElementById('footer-count').textContent =
        `Showing ${Math.min(slice.length, PREVIEW).toLocaleString()} of ${filteredData.length.toLocaleString()} rows`;
    document.getElementById('footer-note').textContent =
        filteredData.length > PREVIEW ? `Preview limited to ${PREVIEW} rows — export for full data` : '';
}

// ── EXPORT ────────────────────────────────────────────
document.getElementById('btn-export').addEventListener('click', async () => {
    if (!finalData.length) return;
    showOverlay('Generating Excel file…');
    try {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Cross-Match Report');
        ws.columns = headers.map(h => ({ header: h, key: h, width: 22 }));
        ws.addRows(finalData);

        // Header row style
        ws.getRow(1).eachCell(cell => {
            cell.font      = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A192F' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });

        const ftthCol = headers.indexOf('FTTH Match') + 1;
        const dslCol  = headers.indexOf('DSL Match')  + 1;

        ws.eachRow((row, rn) => {
            if (rn === 1) return;
            [ftthCol, dslCol].forEach(ci => {
                if (!ci) return;
                const cell = row.getCell(ci);
                if      (cell.value === 'The same')  { cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFD1FAE5'} }; cell.font = { color:{argb:'FF065F46'}, bold:true }; }
                else if (cell.value === 'Changed')   { cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFFEF3C7'} }; cell.font = { color:{argb:'FF92400E'}, bold:true }; }
                else if (cell.value === 'Not avail') { cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFFEE2E2'} }; cell.font = { color:{argb:'FF991B1B'}, bold:true }; }
            });
        });

        const buf = await wb.xlsx.writeBuffer();
        saveAs(new Blob([buf]), 'Samira_CrossMatch_Output.xlsx');
    } finally {
        hideOverlay();
    }
});

// ── OVERLAY HELPERS ───────────────────────────────────
function showOverlay(msg) {
    document.getElementById('overlay-txt').textContent = msg;
    document.getElementById('overlay').classList.add('show');
}
function hideOverlay() {
    document.getElementById('overlay').classList.remove('show');
}
