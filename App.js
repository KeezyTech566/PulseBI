let currentUser = null;
let unifiedDataset = null; // Stores combined rows, file names, and column types
let primaryChartInstance = null;
let secondaryChartInstance = null;

window.addEventListener('DOMContentLoaded', () => {
    const savedUser = localStorage.getItem('pulse_bi_user');
    if (savedUser) {
        currentUser = JSON.parse(savedUser);
        loadUserWorkspace();
        showDashboard();
    }
});

// Auth Flow
function showAuthModal(mode) {
    const isSignup = mode === 'signup';
    document.getElementById('authTitle').innerText = isSignup ? "Create Your Workspace" : "Welcome Back";
    document.getElementById('nameFieldGroup').classList.toggle('hidden', !isSignup);
    document.getElementById('authModal').classList.remove('hidden');
}

function closeAuthModal() { document.getElementById('authModal').classList.add('hidden'); }
function toggleAuthMode() {
    const isSignup = document.getElementById('nameFieldGroup').classList.contains('hidden');
    document.getElementById('authTitle').innerText = isSignup ? "Create Your Workspace" : "Welcome Back";
    document.getElementById('nameFieldGroup').classList.toggle('hidden', !isSignup);
}

function handleAuth(e) {
    e.preventDefault();
    const email = document.getElementById('authEmail').value;
    const name = document.getElementById('authName').value || email.split('@')[0];
    currentUser = { name, email };
    localStorage.setItem('pulse_bi_user', JSON.stringify(currentUser));
    closeAuthModal();
    loadUserWorkspace();
    showDashboard();
}

function logout() {
    localStorage.removeItem('pulse_bi_user');
    currentUser = null;
    document.getElementById('dashboardView').classList.add('hidden');
    document.getElementById('landingView').classList.remove('hidden');
}

function showDashboard() {
    document.getElementById('landingView').classList.add('hidden');
    document.getElementById('dashboardView').classList.remove('hidden');
    document.getElementById('dashboardView').classList.add('flex');
    document.getElementById('userWorkspaceLabel').innerText = `${currentUser.name}'s Analytics Hub`;
}

// Modal Upload Controls
function openUploadModal() {
    document.getElementById('uploadModal').classList.remove('hidden');
    document.getElementById('schemaPreviewArea').classList.add('hidden');
    document.getElementById('uploadLabel').innerText = "Click to browse or drop multiple files here";
}

function closeUploadModal() { document.getElementById('uploadModal').classList.add('hidden'); }

// Multi-File Ingestion Engine
async function handleMultipleFilesSelected(e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    document.getElementById('uploadLabel').innerText = `${files.length} file(s) selected. Processing...`;
    
    let allRows = [];
    let fileNames = [];
    let masterHeadersSet = new Set();

    for (let file of files) {
        fileNames.push(file.name);
        try {
            const rows = await parseFile(file);
            if (rows && rows.length > 0) {
                Object.keys(rows[0]).forEach(h => masterHeadersSet.add(h));
                allRows = allRows.concat(rows);
            }
        } catch (err) {
            console.error(`Error parsing ${file.name}:`, err);
        }
    }

    if (allRows.length === 0) {
        alert("Could not extract valid data rows from the selected files.");
        return;
    }

    const headers = Array.from(masterHeadersSet);
    
    // Normalize rows so all objects share the master header set
    const normalizedRows = allRows.map(row => {
        const newRow = {};
        headers.forEach(h => {
            newRow[h] = row[h] !== undefined ? row[h] : "";
        });
        return newRow;
    });

    // Infer data types across columns
    const columnTypes = {};
    headers.forEach(header => {
        let inferred = 'string';
        for (let i = 0; i < Math.min(normalizedRows.length, 10); i++) {
            const t = inferDataType(normalizedRows[i][header]);
            if (t !== 'string') {
                inferred = t;
                break;
            }
        }
        columnTypes[header] = inferred;
    });

    unifiedDataset = { fileNames, headers, rows: normalizedRows, columnTypes };
    renderSchemaPreviewModal();
}

function parseFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        const fileName = file.name.toLowerCase();

        if (fileName.endsWith('.csv')) {
            reader.onload = (e) => {
                resolve(parseCSVText(e.target.result));
            };
            reader.readAsText(file);
        } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
                    resolve(jsonRows);
                } catch (err) {
                    reject(err);
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            resolve([]);
        }
    });
}

function parseCSVText(text) {
    const lines = text.split('\n').filter(l => l.trim() !== '');
    if (lines.length < 2) return [];
    const headers = lines[0].split(',').map(h => h.trim().replace(/^["']|["']$/g, ''));
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        const currentLine = lines[i].split(',');
        if (currentLine.length > 0) {
            const obj = {};
            headers.forEach((h, index) => {
                let val = currentLine[index] ? currentLine[index].trim().replace(/^["']|["']$/g, '') : "";
                obj[h] = val;
            });
            rows.push(obj);
        }
    }
    return rows;
}

// Data Type Inference
function inferDataType(value) {
    if (value === null || value === undefined || value === '') return 'string';
    if (!isNaN(value)) {
        return Number.isInteger(Number(value)) ? 'integer' : 'decimal';
    }
    const lower = String(value).toLowerCase();
    if (lower === 'true' || lower === 'false') return 'boolean';
    if (!isNaN(Date.parse(value)) && String(value).length > 5) return 'date';
    return 'string';
}

function renderSchemaPreviewModal() {
    const container = document.getElementById('schemaTagsContainer');
    container.innerHTML = '';
    
    unifiedDataset.headers.forEach(h => {
        const type = unifiedDataset.columnTypes[h];
        let badgeColor = "bg-indigo-500/10 text-indigo-400 border-indigo-500/20";
        if (type === 'integer') badgeColor = "bg-emerald-500/10 text-emerald-400 border-emerald-500/20";
        if (type === 'decimal') badgeColor = "bg-amber-500/10 text-amber-400 border-amber-500/20";
        if (type === 'date') badgeColor = "bg-purple-500/10 text-purple-400 border-purple-500/20";

        const tag = document.createElement('div');
        tag.className = `flex items-center justify-between bg-slate-900 border border-slate-800 px-3 py-2 rounded-lg text-xs w-full sm:w-[48%]`;
        tag.innerHTML = `<span class="font-medium text-white truncate mr-2">${h}</span> <span class="border px-2 py-0.5 rounded-md font-mono ${badgeColor}">${type}</span>`;
        container.appendChild(tag);
    });

    document.getElementById('schemaPreviewArea').classList.remove('hidden');
    document.getElementById('uploadLabel').innerText = `${unifiedDataset.fileNames.length} file(s) ready. Schema validated!`;
}

function confirmCreateDashboard() {
    if (!unifiedDataset) return;
    saveUserWorkspace();
    closeUploadModal();
    renderDashboardUI();
}

// Persistence
function saveUserWorkspace() {
    if (!currentUser || !unifiedDataset) return;
    localStorage.setItem(`pulse_bi_multidata_${currentUser.email}`, JSON.stringify(unifiedDataset));
}

function loadUserWorkspace() {
    if (!currentUser) return;
    const stored = localStorage.getItem(`pulse_bi_multidata_${currentUser.email}`);
    if (stored) {
        unifiedDataset = JSON.parse(stored);
        renderDashboardUI();
    }
}

// Render Dashboard & Visualizations
function renderDashboardUI() {
    if (!unifiedDataset) return;

    document.getElementById('emptyNotice').classList.add('hidden');
    document.getElementById('activeDashboardContent').classList.remove('hidden');
    document.getElementById('activeFilesSummary').innerText = `${unifiedDataset.fileNames.length} Files: ${unifiedDataset.fileNames.join(', ')}`;

    // Update KPIs
    document.getElementById('kpiRecords').innerText = unifiedDataset.rows.length.toLocaleString();
    document.getElementById('kpiFileCount').innerText = unifiedDataset.fileNames.length;

    const numericCol = unifiedDataset.headers.find(h => ['integer', 'decimal'].includes(unifiedDataset.columnTypes[h]));
    let metricSum = 0;
    if (numericCol) {
        metricSum = unifiedDataset.rows.reduce((acc, curr) => acc + (parseFloat(curr[numericCol]) || 0), 0);
        document.getElementById('kpiMetricSum').innerText = metricSum > 1000000 ? (metricSum / 1000000).toFixed(2) + 'M' : metricSum.toLocaleString();
    } else {
        document.getElementById('kpiMetricSum').innerText = "N/A";
    }

    // Render Table
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');
    thead.innerHTML = `<tr class="border-b border-slate-800 text-slate-400 uppercase text-xs tracking-wider">${unifiedDataset.headers.map(h => `<th class="py-3 px-4 font-semibold">${h} <span class="text-[10px] text-indigo-400 font-mono lowercase">(${unifiedDataset.columnTypes[h]})</span></th>`).join('')}</tr>`;
    
    tbody.innerHTML = '';
    unifiedDataset.rows.slice(0, 10).forEach(row => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-800/40 transition';
        tr.innerHTML = unifiedDataset.headers.map(h => `<td class="py-3 px-4 text-slate-300 truncate max-w-xs">${row[h]}</td>`).join('');
        tbody.appendChild(tr);
    });

    initCharts(numericCol);
}

// Chart.js Generation
function initCharts(numericCol) {
    const catCol = unifiedDataset.headers.find(h => unifiedDataset.columnTypes[h] === 'string') || unifiedDataset.headers[0];
    
    const counts = {};
    unifiedDataset.rows.forEach(r => {
        const key = r[catCol] || 'Unassigned';
        counts[key] = (counts[key] || 0) + 1;
    });

    const labels = Object.keys(counts).slice(0, 8);
    const dataVals = labels.map(l => counts[l]);

    if (primaryChartInstance) primaryChartInstance.destroy();
    const ctx1 = document.getElementById('primaryChart').getContext('2d');
    primaryChartInstance = new Chart(ctx1, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Frequency Distribution',
                data: dataVals,
                backgroundColor: '#6366f1',
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#94a3b8' } } },
            scales: {
                x: { ticks: { color: '#64748b' }, grid: { color: 'rgba(51, 65, 85, 0.3)' } },
                y: { ticks: { color: '#64748b' }, grid: { color: 'rgba(51, 65, 85, 0.3)' } }
            }
        }
    });

    if (secondaryChartInstance) secondaryChartInstance.destroy();
    const ctx2 = document.getElementById('secondaryChart').getContext('2d');
    secondaryChartInstance = new Chart(ctx2, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: dataVals,
                backgroundColor: ['#6366f1', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#06b6d4'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8', boxWidth: 12 } } },
            cutout: '70%'
        }
    });
}
