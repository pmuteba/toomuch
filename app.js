// --- Dom Elements ---
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const uploadView = document.getElementById('upload-view');
const appView = document.getElementById('app-view');
const toggleUngroupedBtn = document.getElementById('toggle-ungrouped');
const ungroupedPanel = document.getElementById('ungrouped-panel');
const toggleTableBtn = document.getElementById('toggle-table');
const tablePanel = document.getElementById('table-panel');
const treesGrid = document.getElementById('trees-grid');
const ungroupedListEl = document.getElementById('ungrouped-list');

// --- Layout Toggles Variables & Triggers ---
let isUngroupedOpen = true;
let isTableOpen = true;

toggleUngroupedBtn.addEventListener('click', () => {
    isUngroupedOpen = !isUngroupedOpen;
    if (isUngroupedOpen) {
        ungroupedPanel.classList.remove('translate-x-full', 'w-0');
        ungroupedPanel.classList.add('w-80');
        toggleUngroupedBtn.innerText = "⏴";
    } else {
        ungroupedPanel.classList.add('translate-x-full', 'w-0');
        ungroupedPanel.classList.remove('w-80');
        toggleUngroupedBtn.innerText = "⏵";
    }
});

toggleTableBtn.addEventListener('click', () => {
    isTableOpen = !isTableOpen;
    if (isTableOpen) {
        tablePanel.classList.remove('h-10');
        tablePanel.classList.add('h-64');
        toggleTableBtn.innerText = "▼ Row Sheet Data";
    } else {
        tablePanel.classList.remove('h-64');
        tablePanel.classList.add('h-10');
        toggleTableBtn.innerText = "▲ Row Sheet Data";
    }
});

// --- File Drop Listener Logic ---
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('border-indigo-500'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('border-indigo-500'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    const files = e.dataTransfer.files;
    if (files.length) handleExcelParsing(files[0]);
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleExcelParsing(e.target.files[0]);
});

// --- Excel Engine File Parsing ---
function handleExcelParsing(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Transform worksheet data strictly to JSON objects array
        globalData = XLSX.utils.sheet_to_json(worksheet);
        
        // Standardize base column cases if missing
        globalData = globalData.map(row => ({
            'Parent Attribute': row['Parent Attribute'] || '',
            'Attribute': row['Attribute'] || '',
            'Set Status To': row['Set Status To'] || 'WIP',
            'Recommended Merge': row['Recommended Merge'] || '',
            'Other Matches': row['Other Matches'] || '',
            ...row
        }));

        buildStateFromData();
        renderInterface();

        uploadView.classList.add('hidden');
        appView.classList.remove('hidden');
    };
    reader.readAsArrayBuffer(file);
}

// --- Parse Matrix States Logic ---
function buildStateFromData() {
    treeStructures = {};
    ungroupedList = [];

    // Step A: Capture potential unique Parent Merge Roots 
    globalData.forEach(row => {
        if (row['Set Status To'] === 'MERGE' && row['Recommended Merge']) {
            const root = row['Recommended Merge'];
            if (!treeStructures[root]) {
                treeStructures[root] = [];
            }
            treeStructures[root].push(row['Attribute']);
        }
    });

    // Step B: Aggregate any other unique items into Ungrouped
    globalData.forEach(row => {
        if (row['Set Status To'] !== 'MERGE') {
            ungroupedList.push({
                name: row['Attribute'],
                status: row['Set Status To']
            });
        }
    });

    // Enforce Alphabetical sorting constraint on Sidebar items
    ungroupedList.sort((a, b) => a.name.localeCompare(b.name));
}

// --- Component Injections DOM Layout Renderer ---
function renderInterface() {
    renderTreesGrid();
    renderUngroupedSidebar();
    renderTablePreview();
}

function renderTreesGrid() {
    treesGrid.innerHTML = '';
    Object.keys(treeStructures).forEach(rootName => {
        const treeCard = document.createElement('div');
        treeCard.className = 'bg-gray-800 p-4 rounded-xl border border-gray-700 flex flex-col';
        
        // Root element rendering
        let rootHTML = `
            <div class="font-bold text-sm bg-gray-750 p-2 rounded border-b border-indigo-500/30 text-indigo-300 flex items-center justify-between mb-3">
                <span class="truncate">${rootName}</span>
                <span class="text-[10px] uppercase tracking-wider bg-indigo-900/50 text-indigo-400 px-1.5 py-0.5 rounded">Root</span>
            </div>
            <div class="tree-container flex-1 space-y-2" data-root="${rootName}">
        `;

        // Render leaf children array
        treeStructures[rootName].forEach(childName => {
            // Locate row metadata info for Hover triggers
            const rowMeta = globalData.find(r => r['Attribute'] === childName) || {};
            const hoverData = rowMeta['Other Matches'] || '';

            rootHTML += `
                <div class="tree-node text-xs bg-gray-700/50 p-2 rounded border border-gray-600 hover:border-orange-400 relative cursor-grab group" data-attribute="${childName}">
                    <div class="font-medium text-gray-200 truncate">${childName}</div>
                    
                    <!-- Other Matches Inline Custom Popover Context -->
                    ${hoverData ? `
                    <div class="hidden group-hover:block absolute z-30 bg-gray-950 text-gray-300 text-[11px] p-2 rounded shadow-xl border border-gray-700 left-2 right-2 top-full mt-1 max-w-xs">
                        <div class="text-gray-400 font-semibold border-b border-gray-800 pb-1 mb-1">Click to Remap Match:</div>
                        <div class="space-y-1">${parseOtherMatchesHTML(hoverData, childName)}</div>
                    </div>` : ''}
                </div>
            `;
        });

        rootHTML += `</div>`;
        treeCard.innerHTML = rootHTML;
        treesGrid.appendChild(treeCard);
    });
}

function renderUngroupedSidebar() {
    ungroupedListEl.innerHTML = '';
    ungroupedList.forEach(item => {
        const itemEl = document.createElement('div');
        itemEl.className = `text-xs bg-gray-750 p-2.5 rounded border border-gray-700/60 font-medium flex items-center justify-between cursor-grab status-${item.status}`;
        itemEl.setAttribute('data-attribute', item.name);
        itemEl.innerHTML = `
            <span class="truncate pr-1">${item.name}</span>
            <span class="text-[10px] font-bold opacity-80 shrink-0">(${item.status})</span>
        `;
        ungroupedListEl.appendChild(itemEl);
    });
}

// Utility string clean function for Open AI JSON outputs parsing inside sheet cell
function parseOtherMatchesHTML(dictString, currentChild) {
    try {
        // Clean JSON formatting boundaries if string contains literal dictionary text brackets
        let cleanStr = dictString.trim();
        if(cleanStr.startsWith('{') && cleanStr.endsWith('}')) {
            cleanStr = cleanStr.slice(1, -1);
        }
        
        const pairs = cleanStr.split(',');
        return pairs.map(p => {
            const parts = p.split(':');
            if(parts.length < 2) return '';
            const keyName = parts[0].replace(/['"]+/g, '').trim();
            const scoreVal = parts[1].trim();
            
            return `<a href="#" onclick="remapNodeToNewTarget('${currentChild}', '${keyName}'); return false;" class="block text-indigo-400 hover:underline truncate">🔗 ${keyName} (${scoreVal})</a>`;
        }).join('');
    } catch(err) {
        return `<span class="text-red-400">Parsing Match string error</span>`;
    }
}
