// --- Memory State Storage Buffers ---
let globalData = [];      // Raw row JSON data objects
let treeStructures = {};  // Root-to-Leaves hierarchy tracking mapping
let ungroupedList = [];   // Sidebar non-merged data collections array
let draggedAttributeName = null;

// --- Column Filter Storage Matrix ---
let columnFilters = {
    'Parent Attribute': '',
    'Attribute': '',
    'Set Status To': '',
    'Recommended Merge': ''
};

// --- DOM Layout Element Selectors ---
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

// --- Layout Collapse Panel Toggle Triggers ---
let isUngroupedOpen = true;
let isTableOpen = true;

toggleUngroupedBtn.addEventListener('click', () => {
    isUngroupedOpen = !isUngroupedOpen;
    if (isUngroupedOpen) {
        ungroupedPanel.classList.remove('sidebar-hidden');
        toggleUngroupedBtn.innerText = "⏴";
    } else {
        ungroupedPanel.classList.add('sidebar-hidden');
        toggleUngroupedBtn.innerText = "⏵";
    }
});

toggleTableBtn.addEventListener('click', () => {
    isTableOpen = !isTableOpen;
    if (isTableOpen) {
        tablePanel.classList.remove('footer-collapsed');
        toggleTableBtn.innerText = "▼ Row Sheet Data";
    } else {
        tablePanel.classList.add('footer-collapsed');
        toggleTableBtn.innerText = "▲ Row Sheet Data";
    }
});

// --- File Selection Event Handlers ---
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => { 
    e.preventDefault(); 
    dropZone.classList.add('border-indigo-500'); 
});

dropZone.addEventListener('dragleave', () => dropZone.classList.remove('border-indigo-500'));

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('border-indigo-500');
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
        handleExcelParsing(files[0]); // Target explicit file array entry index
    }
});

fileInput.addEventListener('change', (e) => {
    const files = e.target.files;
    if (files && files.length > 0) {
        handleExcelParsing(files[0]); // Target explicit file array entry index
    }
});

// --- Excel Workbook Parsing Engine ---
function handleExcelParsing(fileObject) {
    if (!fileObject) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const firstSheetName = workbook.SheetNames[0]; // Fix index accessor
            const worksheet = workbook.Sheets[firstSheetName];
            
            globalData = XLSX.utils.sheet_to_json(worksheet);
            
            if (!globalData || globalData.length === 0) {
                alert("The first sheet appears to be empty.");
                return;
            }

            // Standardize cell values across casing and unexpected spaces
            globalData = globalData.map(row => {
                let rParent = row['Parent Attribute'] || '';
                let rAttr = row['Attribute'] || '';
                let rStatus = row['Set Status To'] || 'WIP';
                let rMerge = row['Recommended Merge'] || '';
                let rMatches = row['Other Matches'] || '';

                return {
                    ...row,
                    'Parent Attribute': String(rParent).trim(),
                    'Attribute': String(rAttr).trim(),
                    'Set Status To': String(rStatus).trim().toUpperCase(),
                    'Recommended Merge': String(rMerge).trim(),
                    'Other Matches': String(rMatches).trim()
                };
            });

            buildStateFromData();
            renderInterface();

            uploadView.classList.add('hidden');
            appView.classList.remove('hidden');
            
        } catch (error) {
            console.error("Excel processing crash:", error);
            alert("Error parsing Excel content. Check developer console (F12) for logs.");
        }
    };
    reader.readAsArrayBuffer(fileObject);
}

// --- Parse Matrix States Logic ---
function buildStateFromData() {
    treeStructures = {};
    ungroupedList = [];

    // Map out unique parent merge roots
    globalData.forEach(row => {
        if (row['Set Status To'] === 'MERGE' && row['Recommended Merge']) {
            const root = row['Recommended Merge'];
            if (!treeStructures[root]) {
                treeStructures[root] = [];
            }
            if (row['Attribute'] && !treeStructures[root].includes(row['Attribute'])) {
                treeStructures[root].push(row['Attribute']);
            }
        }
    });

    // Populate sidebar array with ungrouped items
    globalData.forEach(row => {
        if (row['Set Status To'] !== 'MERGE' && row['Attribute']) {
            // Avoid duplicates in list
            if (!ungroupedList.some(item => item.name === row['Attribute'])) {
                ungroupedList.push({
                    name: row['Attribute'],
                    status: row['Set Status To']
                });
            }
        }
    });

    ungroupedList.sort((a, b) => a.name.localeCompare(b.name));
}

// --- Component UI Renderer Pipeline ---
function renderInterface() {
    renderTreesGrid();
    renderUngroupedSidebar();
    renderTablePreview();
    initializeDragAndDropListeners();
}

function renderTreesGrid() {
    treesGrid.innerHTML = '';
    const roots = Object.keys(treeStructures);

    if (roots.length === 0) {
        treesGrid.innerHTML = `<div class="text-gray-500 italic p-4 col-span-full">No active merge hierarchies mapped yet. Drag items here to spawn trees.</div>`;
        return;
    }

    roots.forEach(rootName => {
        const treeCard = document.createElement('div');
        treeCard.className = 'tree-card';
        
        let rootHTML = `
            <div class="font-bold text-sm bg-gray-700/50 p-2 rounded border-b border-indigo-500/30 text-indigo-300 flex items-center justify-between mb-2">
                <span class="truncate" title="${rootName}">${rootName}</span>
                <span class="text-[10px] uppercase tracking-wider bg-indigo-900/50 text-indigo-400 px-1.5 py-0.5 rounded">Root</span>
            </div>
            <div class="tree-container" data-root="${rootName}">
        `;

        treeStructures[rootName].forEach(childName => {
            const rowMeta = globalData.find(r => r['Attribute'] === childName) || {};
            const hoverData = rowMeta['Other Matches'] || '';

            rootHTML += `
                <div class="tree-node relative group" data-attribute="${childName}">
                    <div class="text-xs font-medium text-gray-200 truncate">${childName}</div>
                    
                    ${hoverData ? `
                    <div class="hidden group-hover:block absolute left-0 right-0 top-full mt-1 bg-gray-950 text-gray-300 text-[11px] p-2 rounded shadow-2xl border border-gray-700 z-30 max-w-xs">
                        <div class="text-gray-400 font-semibold border-b border-gray-800 pb-1 mb-1">Remap Target:</div>
                        <div class="space-y-1 max-h-32 overflow-y-auto">${parseOtherMatchesHTML(hoverData, childName)}</div>
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
    if (ungroupedList.length === 0) {
        ungroupedListEl.innerHTML = `<div class="text-gray-500 italic text-center p-4">Empty list.</div>`;
        return;
    }

    ungroupedList.forEach(item => {
        const itemEl = document.createElement('div');
        itemEl.className = `text-xs p-2.5 rounded border border-gray-700/60 font-medium flex items-center justify-between status-${item.status}`;
        itemEl.setAttribute('data-attribute', item.name);
        itemEl.innerHTML = `
            <span class="truncate pr-1">${item.name}</span>
            <span class="text-[10px] font-bold opacity-80 shrink-0">(${item.status})</span>
        `;
        ungroupedListEl.appendChild(itemEl);
    });
}

function renderTablePreview() {
    const tableEl = document.getElementById('data-table');
    tableEl.innerHTML = '';

    if (globalData.length === 0) return;

    const headers = ['Parent Attribute', 'Attribute', 'Set Status To', 'Recommended Merge', 'Other Matches'];

    let headerHTML = `<thead class="bg-gray-850 sticky top-0 border-b border-gray-700 z-10"><tr>`;
    headers.forEach(h => {
        headerHTML += `
            <th class="p-2 text-gray-400 font-semibold border-r border-gray-700 min-w-[160px]">
                <div class="flex flex-col space-y-1">
                    <span>${h}</span>
                    ${h !== 'Other Matches' ? `
                    <input type="text" placeholder="Filter..." value="${columnFilters[h] || ''}" oninput="updateColumnFilter('${h}', this.value)" class="w-full bg-gray-900 border border-gray-700 rounded px-1.5 py-0.5 text-[11px] text-gray-200 focus:outline-none focus:border-indigo-500">` : ''}
                </div>
            </th>`;
    });
    headerHTML += `</tr></thead>`;

    let bodyHTML = `<tbody>`;
    globalData.forEach((row, rowIndex) => {
        // Run column string match filters
        if (columnFilters('Parent Attribute') && !String(row('Parent Attribute') || '').toLowerCase().includes(columnFilters('Parent Attribute').toLowerCase())) return;
        if (columnFilters('Attribute') && !String(row('Attribute') || '').toLowerCase().includes(columnFilters('Attribute').toLowerCase())) return;
        if (columnFilters('Set Status To') && !String(row('Set Status To') || '').toLowerCase().includes(columnFilters('Set Status To').toLowerCase())) return;
        if (columnFilters('Recommended Merge') && !String(row('Recommended Merge') || '').toLowerCase().includes(columnFilters('Recommended Merge').toLowerCase())) return;
        bodyHTML += <tr class="border-b border-gray-800 hover:bg-gray-750/30">;
        headers.forEach(h => {
            const cellValue = row(h) || '';
            if (h === 'Other Matches') {
                bodyHTML += <td class="p-2 border-r border-gray-800 text-gray-500 truncate max-w-xs" title="${cellValue}">${cellValue}</td>;
                    } 
            else if (h === 'Set Status To') {
                bodyHTML +=  <td class="p-1 border-r border-gray-800"> <select onchange="updateTableCell(${rowIndex}, '${h}', this.value)" class="w-full bg-gray-900 border border-gray-700 rounded px-1 py-0.5 text-xs text-gray-200 focus:outline-none focus:border-indigo-500"> <option value="MERGE" ${cellValue === 'MERGE' ? 'selected' : ''}>MERGE</option> <option value="NAV" ${cellValue === 'NAV' ? 'selected' : ''}>NAV</option> <option value="KEEP" ${cellValue === 'KEEP' ? 'selected' : ''}>KEEP</option> <option value="DROP" ${cellValue === 'DROP' ? 'selected' : ''}>DROP</option> <option value="WIP" ${cellValue === 'WIP' ? 'selected' : ''}>WIP</option> </select> </td>;
                    } 
            else {
                bodyHTML +=  <td class="p-1 border-r border-gray-800"> <input type="text" value="${cellValue}" onchange="updateTableCell(${rowIndex}, '${h}', this.value)" class="w-full bg-transparent px-2 py-1 text-xs text-gray-200 focus:bg-gray-900 focus:outline-none rounded"> </td>;
                    }
        });
        bodyHTML += </tr>;
            });
    bodyHTML += </tbody>;
        tableEl.innerHTML = headerHTML + bodyHTML;
}
// --- Text Utility Parsers ---
function parseOtherMatchesHTML(dictString, currentChild) {
    try {
        let str = dictString.trim();
        if(str.startsWith('{') && str.endsWith('}')) str = str.slice(1, -1);
        if(!str) return 'None';
        const pairs = str.split(',');
        return pairs.map(p => {
            const parts = p.split(':');
            if(parts.length < 2) return '';
            const keyName = parts(0).replace(/('")+/g, '').trim();const scoreVal = parts(1).trim();
            // Build absolute scope global executable functions link anchor
            return <a href="#" onclick="remapNodeToNewTarget('${currentChild}', '${keyName}'); return false;" class="block text-indigo-400 hover:underline truncate">🔗 ${keyName} (${scoreVal})</a>;
            }).join('');
} catch(err) {
    return <span class="text-red-400">Parsing Error</span>;
    }
}
function remapNodeToNewTarget(childAttribute, targetMatchParent) {
    let matchedRow = globalData.find(r => r('Attribute') === childAttribute);
    if (!matchedRow) return;
    matchedRow('Recommended Merge') = targetMatchParent;
    matchedRow('Set Status To') = 'MERGE';
    buildStateFromData();renderInterface();
}
function updateTableCell(rowIndex, columnName, value) {
    globalData(rowIndex)(columnName) = value;
}
function updateColumnFilter(columnName, value) {
    columnFilters(columnName) = value;
    renderTablePreview();
}
// --- Drag and Drop API Engine Controllers ---
function initializeDragAndDropListeners() {
    const draggableItems = document.querySelectorAll('(data-attribute)');
    draggableItems.forEach(item => {
        item.setAttribute('draggable', 'true');
        item.addEventListener('dragstart', (e) => {
            draggedAttributeName = e.target.getAttribute('data-attribute');
            e.dataTransfer.setData('text/plain', draggedAttributeName);
            e.target.classList.add('opacity-30');
        });
        item.addEventListener('dragend', (e) => {
            e.target.classList.remove('opacity-30');
        });
    });
    const containers = document.querySelectorAll('.tree-container');
    containers.forEach(c => {
        c.addEventListener('dragover', (e) => e.preventDefault());
        c.addEventListener('drop', (e) => {
            e.preventDefault();
            const targetRoot = c.getAttribute('data-root');
            if (draggedAttributeName && draggedAttributeName !== targetRoot) {
                let row = globalData.find(r => r('Attribute') === draggedAttributeName);
                if (row) {
                    row('Recommended Merge') = targetRoot;
                    row('Set Status To') = 'MERGE';
                    syncAndRebuildView();
                }
            }
        });
    });
    const mainCanvas = document.getElementById('trees-window');
    mainCanvas.addEventListener('dragover', (e) => e.preventDefault());
    mainCanvas.addEventListener('drop', (e) => {
        if (e.target === mainCanvas || e.target === document.getElementById('trees-grid')) {
            e.preventDefault();
            if (draggedAttributeName) {
                // Spawn a new root node if it was a childlet 
                row = globalData.find(r => r('Attribute') === draggedAttributeName);
                if (row) {
                    row('Set Status To') = 'WIP';
                    row('Recommended Merge') = '';
                    syncAndRebuildView();
                }
            }
        }
    });
    ungroupedPanel.addEventListener('dragover', (e) => e.preventDefault());
    ungroupedPanel.addEventListener('drop', (e) => {
        e.preventDefault();
        if (draggedAttributeName) {
            // Is it currently an active tree parent root element?
            if (Object.keys(treeStructures).includes(draggedAttributeName)) {
                globalData.forEach(row => {
                    if (row('Recommended Merge') === draggedAttributeName) {
                        row('Set Status To') = 'WIP';row('Recommended Merge') = '';
                    }
                });
            }
            let row = globalData.find(r => r('Attribute') === draggedAttributeName);
            if (row) {
                row('Set Status To') = 'WIP';
                row('Recommended Merge') = '';
            }
            syncAndRebuildView();
        }
    });
}
function syncAndRebuildView() {
    draggedAttributeName = null;
    buildStateFromData();
    renderInterface();
}
// --- Global Synchronization Event Action Buttons ---
document.getElementById('btn-refresh').addEventListener('click', () => {
    buildStateFromData();
    renderInterface();
});
document.getElementById('btn-save').addEventListener('click', () => {
    globalData.forEach(row => {
        const currentAttr = row('Attribute');
        let foundInTree = false;
        Object.keys(treeStructures).forEach(rootName => {if (treeStructures(rootName).includes(currentAttr)) {
            row('Set Status To') = 'MERGE';
            row('Recommended Merge') = rootName;foundInTree = true;
        }
                                                        });
        if (Object.keys(treeStructures).includes(currentAttr)) {
            row('Recommended Merge') = '';
            if (treeStructures(currentAttr).length === 0 && row('Set Status To') === 'MERGE') {
                row('Set Status To') = 'WIP';
            }
            foundInTree = true;
        }
        if (!foundInTree) {
            const sideMatch = ungroupedList.find(item => item.name === currentAttr);
            if (sideMatch) {
                row('Set Status To') = sideMatch.status;
                row('Recommended Merge') = '';
            }
        }
    });
    renderTablePreview();
    alert('Dashboard state successfully saved and synchronized into the spreadsheet layout.');
});
document.getElementById('btn-download').addEventListener('click', () => {
    if (globalData.length === 0) {
        alert("No data array available to download.");
        return;
    }
    // Set Parent Attribute values for child columns prior to download
    globalData.forEach(row => {
        const attrName = row('Attribute');
        if (treeStructures(attrName) && treeStructures(attrName).length > 0) {
            treeStructures(attrName).forEach(childName => {
                let childRow = globalData.find(r => r('Attribute') === childName);
                if (childRow) {
                    childRow('Parent Attribute') = attrName;
                }
            });
        }
    });
    const worksheet = XLSX.utils.json_to_sheet(globalData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Master Attributes Data");
    XLSX.writeFile(workbook, "Deduplicated_Attributes_Master.xlsx");
});
    
