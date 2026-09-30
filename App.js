let currentUser = null;
let activeDataset = null; 
let customMeasures = {};

let dashboardPages = [
    { id: 'page-1', name: 'Executive Overview', visuals: [] },
    { id: 'page-2', name: 'Operations & Safety', visuals: [] }
];
let activePageId = 'page-1';

window.addEventListener('DOMContentLoaded', () => {
    const savedUser = localStorage.getItem('pulse_bi_user_v9');
    if (savedUser) {
        currentUser = JSON.parse(savedUser);
        showSetupView();
    } else {
        showOnboardingView();
    }
});

function showOnboardingView() {
    document.getElementById('onboardingView').classList.remove('hidden');
    document.getElementById('setupView').classList.add('hidden');
    document.getElementById('dashboardView').classList.add('hidden');
}

function showSetupView() {
    document.getElementById('onboardingView').classList.add('hidden');
    document.getElementById('setupView').classList.remove('hidden');
    document.getElementById('setupView').classList.add('flex');
    document.getElementById('dashboardView').classList.add('hidden');
    if (currentUser) {
        document.getElementById('setupUserSubtitle').innerText = `${currentUser.name} (${currentUser.role})`;
    }
}

function showDashboardView() {
    document.getElementById('onboardingView').classList.add('hidden');
    document.getElementById('setupView').classList.add('hidden');
    document.getElementById('dashboardView').classList.remove('hidden');
    document.getElementById('dashboardView').classList.add('flex');
    
    document.getElementById('roleBadge').innerText = `${currentUser.role} Portal`;
    document.getElementById('userWorkspaceLabel').innerText = `${currentUser.name}`;
}

function goToSetupView() { showSetupView(); }

function showAuthModal(mode) { document.getElementById('authModal').classList.remove('hidden'); }
function closeAuthModal() { document.getElementById('authModal').classList.add('hidden'); }

function handleAuth(e) {
    e.preventDefault();
    const email = document.getElementById('authEmail').value;
    const name = document.getElementById('authName').value || email.split('@')[0];
    const role = document.getElementById('userRoleSelect').value;
    
    currentUser = { name, email, role };
    localStorage.setItem('pulse_bi_user_v9', JSON.stringify(currentUser));
    closeAuthModal();
    showSetupView();
}

function logout() {
    localStorage.removeItem('pulse_bi_user_v9');
    currentUser = null;
    showOnboardingView();
}

async function handleSetupFilesSelected(e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;
    const rows = await parseFile(files[0]);
    const headers = Object.keys(rows[0] || {});
    const columnTypes = {};
    headers.forEach(h => columnTypes[h] = 'string');

    activeDataset = { name: files[0].name, headers, rows, columnTypes };
    showDashboardView();
    initStudio();
}

function parseFile(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const workbook = XLSX.read(e.target.result, { type: 'array' });
            const jsonRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });
            resolve(jsonRows);
        };
        reader.readAsArrayBuffer(file);
    });
}

function initStudio() {
    if (!activeDataset) return;

    const xSelect = document.getElementById('visualXAxis');
    const ySelect = document.getElementById('visualYAxis');
    xSelect.innerHTML = '';
    ySelect.innerHTML = '';

    activeDataset.headers.forEach(h => {
        xSelect.innerHTML += `<option value="${h}">${h}</option>`;
        ySelect.innerHTML += `<option value="${h}">${h}</option>`;
    });

    if (activeDataset.headers.length > 1) {
        xSelect.value = activeDataset.headers[0];
        ySelect.value = activeDataset.headers[1];
    }

    renderPageTabs();
    renderActivePageCanvas();
    renderTableRecords(activeDataset.rows);
}

function renderPageTabs() {
    const container = document.getElementById('pagesTabContainer');
    container.innerHTML = '';

    dashboardPages.forEach(page => {
        const isActive = page.id === activePageId;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = `px-4 py-2 rounded-lg text-xs font-semibold transition cursor-pointer whitespace-nowrap border ${isActive ? 'bg-indigo-600 text-white border-indigo-500 shadow-lg' : 'bg-slate-800 text-slate-300 border-slate-700 hover:bg-slate-700'}`;
        btn.innerHTML = `<i class="fa-solid fa-file-lines mr-1.5"></i> ${page.name}`;
        btn.onclick = () => switchPage(page.id);
        container.appendChild(btn);
    });
}

function switchPage(pageId) {
    activePageId = pageId;
    renderPageTabs();
    renderActivePageCanvas();
}

function addNewDashboardPage() {
    const name = prompt("Enter a name for the new dashboard page:", `Page ${dashboardPages.length + 1}`);
    if (name) {
        const newId = 'page-' + Date.now();
        dashboardPages.push({ id: newId, name: name.trim(), visuals: [] });
        switchPage(newId);
    }
}

function renameCurrentPage() {
    const page = dashboardPages.find(p => p.id === activePageId);
    if (!page) return;
    const newName = prompt("Rename current dashboard page:", page.name);
    if (newName) {
        page.name = newName.trim();
        renderPageTabs();
        renderActivePageCanvas();
    }
}

function addVisualToCurrentPage() {
    if (!activeDataset) {
        alert("Please upload a dataset first in Environment Setup.");
        return;
    }

    const page = dashboardPages.find(p => p.id === activePageId);
    if (!page) return;

    const visualConfig = {
        id: 'vis-' + Date.now(),
        type: document.getElementById('visualTypeSelect').value,
        xAxis: document.getElementById('visualXAxis').value,
        yAxis: document.getElementById('visualYAxis').value,
        agg: document.getElementById('visualAgg').value,
        startDate: document.getElementById('calendarStart').value,
        endDate: document.getElementById('calendarEnd').value
    };

    page.visuals.push(visualConfig);
    renderActivePageCanvas();
}

function removeVisualFromPage(visId) {
    const page = dashboardPages.find(p => p.id === activePageId);
    if (!page) return;
    page.visuals = page.visuals.filter(v => v.id !== visId);
    renderActivePageCanvas();
}

function renderActivePageCanvas() {
    const page = dashboardPages.find(p => p.id === activePageId);
    if (!page) return;

    document.getElementById('currentPageTitleHeading').innerText = `Canvas Studio: ${page.name}`;
    const container = document.getElementById('pageVisualsContainer');
    container.innerHTML = '';

    if (page.visuals.length === 0) {
        container.innerHTML = `
            <div class="col-span-2 bg-slate-900 border border-dashed border-slate-700 rounded-2xl p-12 text-center space-y-3">
                <div class="text-indigo-400 text-3xl"><i class="fa-solid fa-chart-column"></i></div>
                <h3 class="font-bold text-white text-base">Canvas Page is Empty</h3>
                <p class="text-xs text-slate-400 max-w-sm mx-auto">Use the Visual Builder pane above to configure and pin multiple charts to this page.</p>
            </div>
        `;
        return;
    }

    page.visuals.forEach((vis) => {
        const canvasId = `canvas-chart-${vis.id}`;

        const card = document.createElement('div');
        card.className = "bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-xl flex flex-col justify-between space-y-4";
        card.innerHTML = `
            <div class="flex items-center justify-between">
                <div>
                    <h4 class="text-sm font-bold text-white">${vis.type.toUpperCase()}: ${vis.yAxis} by ${vis.xAxis}</h4>
                    <p class="text-[11px] text-slate-400">Measure / Aggregation: ${vis.agg.toUpperCase()}</p>
                </div>
                <button type="button" onclick="removeVisualFromPage('${vis.id}')" class="text-slate-500 hover:text-rose-400 text-xs cursor-pointer p-1" title="Remove Visual"><i class="fa-solid fa-trash"></i></button>
            </div>
            <div class="relative w-full h-72 flex items-center justify-center">
                <canvas id="${canvasId}"></canvas>
            </div>
        `;
        container.appendChild(card);

        setTimeout(() => renderChartOnCanvas(vis, canvasId), 50);
    });
}

function renderChartOnCanvas(vis, canvasId) {
    const ctxElem = document.getElementById(canvasId);
    if (!ctxElem) return;

    let filteredRows = activeDataset.rows;
    if (vis.startDate || vis.endDate) {
        filteredRows = activeDataset.rows.filter(r => {
            const rowDate = r.date || r.created_at || r.month || '';
            if (!rowDate) return true;
            if (vis.startDate && rowDate < vis.startDate) return false;
            if (vis.endDate && rowDate > vis.endDate) return false;
            return true;
        });
    }

    // Expanded Aggregation Computation Engine
    const aggregated = {};
    filteredRows.forEach(r => {
        const key = String(r[vis.xAxis] || 'Unassigned');
        const val = parseFloat(r[vis.yAxis]) || 0;
        if (!aggregated[key]) aggregated[key] = { values: [], sum: 0, count: 0, min: Infinity, max: -Infinity };
        aggregated[key].values.push(val);
        aggregated[key].sum += val;
        aggregated[key].count += 1;
        if (val < aggregated[key].min) aggregated[key].min = val;
        if (val > aggregated[key].max) aggregated[key].max = val;
    });

    const labels = Object.keys(aggregated);
    const dataVals = labels.map(l => {
        const group = aggregated[l];
        const agg = vis.agg;
        if (agg === 'avg') return group.sum / group.count;
        if (agg === 'count' || agg === 'counta') return group.count;
        if (agg === 'distinctcount') return new Set(group.values).size;
        if (agg === 'min') return group.min === Infinity ? 0 : group.min;
        if (agg === 'max') return group.max === -Infinity ? 0 : group.max;
        if (agg === 'median') {
            const sorted = [...group.values].sort((a,b) => a - b);
            const mid = Math.floor(sorted.length / 2);
            return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
        }
        if (agg === 'variance' || agg === 'stddev') {
            const mean = group.sum / group.count;
            const squareDiffs = group.values.map(v => Math.pow(v - mean, 2));
            const avgSquareDiff = squareDiffs.reduce((a, b) => a + b, 0) / group.count;
            return agg === 'stddev' ? Math.sqrt(avgSquareDiff) : avgSquareDiff;
        }
        return group.sum; // default sum
    });

    let chartConfigType = 'bar';
    let fillArea = false;
    if (vis.type === 'line') chartConfigType = 'line';
    if (vis.type === 'area') { chartConfigType = 'line'; fillArea = true; }
    if (vis.type === 'radar') chartConfigType = 'radar';
    if (vis.type === 'polar') chartConfigType = 'polarArea';
    if (vis.type === 'doughnut') chartConfigType = 'doughnut';

    new Chart(ctxElem.getContext('2d'), {
        type: chartConfigType,
        data: {
            labels: labels,
            datasets: [{
                label: `${vis.yAxis} (${vis.agg})`,
                data: dataVals,
                backgroundColor: ['#6366f1', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#06b6d4'],
                borderColor: '#818cf8',
                borderWidth: chartConfigType === 'line' || chartConfigType === 'area' ? 3 : 1,
                fill: fillArea,
                borderRadius: 6,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: ['doughnut', 'polarArea', 'radar'].includes(chartConfigType), position: 'bottom', labels: { color: '#94a3b8' } } },
            scales: ['doughnut', 'polarArea'].includes(chartConfigType) ? {} : {
                x: { ticks: { color: '#64748b', font: { size: 9 } }, grid: { color: 'rgba(51, 65, 85, 0.3)' } },
                y: { ticks: { color: '#64748b', font: { size: 9 } }, grid: { color: 'rgba(51, 65, 85, 0.3)' } }
            }
        }
    });
}

function renderTableRecords(rows) {
    if (!activeDataset) return;
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');

    if (rows.length === 0) {
        thead.innerHTML = '';
        tbody.innerHTML = `<tr><td colspan="5" class="py-6 text-center text-slate-500">No data records found. Upload files in Setup to populate table.</td></tr>`;
        return;
    }

    thead.innerHTML = `<tr class="border-b border-slate-800 text-slate-400 uppercase">${activeDataset.headers.map(h => `<th class="py-2 px-3">${h}</th>`).join('')}</tr>`;
    tbody.innerHTML = '';
    rows.forEach(r => {
        tbody.innerHTML += `<tr class="hover:bg-slate-800/40">${activeDataset.headers.map(h => `<td class="py-2 px-3">${r[h]}</td>`).join('')}</tr>`;
    });
}

function searchTableRecords() {
    if (!activeDataset) return;
    const query = document.getElementById('tableSearch').value.toLowerCase();
    const filtered = activeDataset.rows.filter(r => Object.values(r).some(v => String(v).toLowerCase().includes(query)));
    renderTableRecords(filtered);
}

function openMeasureModal() { document.getElementById('measureModal').classList.remove('hidden'); }
function closeMeasureModal() { document.getElementById('measureModal').classList.add('hidden'); }
function saveNewMeasure() {
    const name = document.getElementById('measureName').value;
    const expr = document.getElementById('measureExpr').value;
    if (name && expr) {
        customMeasures[name] = expr;
        alert(`Measure "${name}" successfully authored and compiled!`);
        closeMeasureModal();
    }
}

function openPublishModal() { document.getElementById('publishModal').classList.remove('hidden'); }
function closePublishModal() { document.getElementById('publishModal').classList.add('hidden'); }

function sendReportToEmail() {
    const provider = document.getElementById('emailProviderSelect').value;
    const email = document.getElementById('publishEmailInput').value;
    const note = document.getElementById('publishNoteInput').value || 'Please review the attached compiled PulseBI analytical report.';

    if (!email) {
        alert("Please enter a valid recipient email address.");
        return;
    }

    const subject = encodeURIComponent(`PulseBI Report Suite: ${currentUser ? currentUser.role : 'Executive'} Summary`);
    const body = encodeURIComponent(`Recipient: ${email}\n\nSender: ${currentUser ? currentUser.name : 'User'}\nRole: ${currentUser ? currentUser.role : 'Analyst'}\nPages: ${dashboardPages.map(p => p.name).join(', ')}\n\nMessage:\n${note}\n\n--- Generated via PulseBI Enterprise Studio ---`);

    let mailtoLink = `mailto:${email}?subject=${subject}&body=${body}`;

    if (provider === 'gmail') {
        mailtoLink = `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(email)}&su=${subject}&body=${body}`;
        window.open(mailtoLink, '_blank');
    } else if (provider === 'outlook') {
        mailtoLink = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(email)}&subject=${subject}&body=${body}`;
        window.open(mailtoLink, '_blank');
    } else {
        window.location.href = mailtoLink;
    }

    alert(`Report prepared! Opening ${provider.toUpperCase()} to send to ${email}.`);
    closePublishModal();
}
