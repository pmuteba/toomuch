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
function handleExcelParsing(fileList) {
    const file = fileList[0]; // Safely pull the first file item from the drop array
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // FIXED: Target the first sheet index [0] explicitly from the sheet names array
            const firstSheetName = workbook.SheetNames[0]; 
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Transform worksheet data strictly to JSON objects array
            globalData = XLSX.utils.sheet_to_json(worksheet);
            
            if (!globalData || globalData.length === 0) {
                alert("The first sheet appears to be empty.");
                return;
            }

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

            // Transition UI views
            uploadView.classList.add('hidden');
            appView.classList.remove('hidden');
            
        } catch (error) {
            console.error("Parsing Error details:", error);
            alert("Failed to parse the Excel file. Please open your browser console (F12) to see details.");
        }
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
    
    // Inject and connect HTML5 state listener loops to DOM elements
    initializeDragAndDropListeners();
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

// --- Interactive Click Node Relocation Engine ---
function remapNodeToNewTarget(childAttribute, targetMatchParent) {
    // Locate the row inside our main state array
    let matchedRow = globalData.find(r => r['Attribute'] === childAttribute);
    if (!matchedRow) return;

    // Check if the target root parent exists in our current active tree layout
    let targetExistsInTrees = Object.keys(treeStructures).includes(targetMatchParent);
    let targetExistsInUngrouped = ungroupedList.some(item => item.name === targetMatchParent);

    if (targetExistsInTrees) {
        // Shift child over to target tree structure
        matchedRow['Recommended Merge'] = targetMatchParent;
        matchedRow['Set Status To'] = 'MERGE';
    } else if (targetExistsInUngrouped) {
        // Convert the target match attribute into a new Root element
        let parentRow = globalData.find(r => r['Attribute'] === targetMatchParent);
        if (parentRow) parentRow['Set Status To'] = 'WIP'; // Becomes root tracking container
        
        matchedRow['Recommended Merge'] = targetMatchParent;
        matchedRow['Set Status To'] = 'MERGE';
    } else {
        // Fallback: Build a brand new custom root node tracking point
        matchedRow['Recommended Merge'] = targetMatchParent;
        matchedRow['Set Status To'] = 'MERGE';
    }

    // Refresh application boundaries and update downstream UI sheets
    buildStateFromData();
    renderInterface();
}

// --- Drag & Drop Native API Event Handlers ---
let draggedAttributeName = null;

function initializeDragAndDropListeners() {
    // Select all draggable nodes inside both windows
    const draggableItems = document.querySelectorAll('[data-attribute]');
    draggableItems.forEach(item => {
        item.setAttribute('draggable', 'true');
        
        item.addEventListener('dragstart', (e) => {
            draggedAttributeName = e.target.getAttribute('data-attribute');
            e.dataTransfer.setData('text/plain', draggedAttributeName);
            e.target.classList.add('opacity-40');
        });

        item.addEventListener('dragend', (e) => {
            e.target.classList.remove('opacity-40');
        });
    });

    // Setup Target A: The Main Trees Canvas View Window area
    const treesContainerEls = document.querySelectorAll('.tree-container');
    treesContainerEls.forEach(container => {
        container.addEventListener('dragover', (e) => e.preventDefault());
        container.addEventListener('drop', (e) => {
            e.preventDefault();
            const targetRoot = container.getAttribute('data-root');
            handleNodeDropOnTree(draggedAttributeName, targetRoot);
        });
    });

    // Allow dropping on empty space in the Trees window to spawn new roots
    const mainTreesWindow = document.getElementById('trees-window');
    mainTreesWindow.addEventListener('dragover', (e) => e.preventDefault());
    mainTreesWindow.addEventListener('drop', (e) => {
        // Ensure we are dropping directly on the blank container backdrop canvas
        if (e.target === mainTreesWindow || e.target === document.getElementById('trees-grid')) {
            e.preventDefault();
            handleNodeDropOnEmptyCanvas(draggedAttributeName);
        }
    });

    // Setup Target B: The Sidebar Ungrouped panel boundary
    ungroupedPanel.addEventListener('dragover', (e) => e.preventDefault());
    ungroupedPanel.addEventListener('drop', (e) => {
        e.preventDefault();
        handleNodeDropOnUngrouped(draggedAttributeName);
    });
}

// Logic: Moving attributes over into an active tree structure
function handleNodeDropOnTree(attrName, targetRoot) {
    if (!attrName || attrName === targetRoot) return;

    let row = globalData.find(r => r['Attribute'] === attrName);
    if (!row) return;

    row['Recommended Merge'] = targetRoot;
    row['Set Status To'] = 'MERGE';

    syncAndRebuildView();
}

// Logic: Dropping an attribute onto empty canvas space to spawn a new tree root
function handleNodeDropOnEmptyCanvas(attrName) {
    if (!attrName) return;

    // Check if the attribute is currently an active tree root
    const isCurrentRoot = Object.keys(treeStructures).includes(attrName);

    if (isCurrentRoot) {
        // Detach and ungroup its children
        globalData.forEach(row => {
            if (row['Recommended Merge'] === attrName && row['Set Status To'] === 'MERGE') {
                row['Set Status To'] = 'WIP'; // Default status
                row['Recommended Merge'] = '';
            }
        });
    } else {
        // Convert child element or ungrouped item into a root node
        let row = globalData.find(r => r['Attribute'] === attrName);
        if (row) {
            row['Set Status To'] = 'WIP'; // Set root status to WIP by default
            row['Recommended Merge'] = '';
        }
    }

    syncAndRebuildView();
}

// Logic: Moving attributes out into the sidebar list panel
function handleNodeDropOnUngrouped(attrName) {
    if (!attrName) return;

    // If it's a root node, ungroup its children along with it
    const isRoot = Object.keys(treeStructures).includes(attrName);
    if (isRoot) {
        globalData.forEach(row => {
            if (row['Recommended Merge'] === attrName) {
                row['Set Status To'] = 'WIP'; // Default status
                row['Recommended Merge'] = '';
            }
        });
    }

    // Reset row status to WIP
    let row = globalData.find(r => r['Attribute'] === attrName);
    if (row) {
        row['Set Status To'] = 'WIP';
        row['Recommended Merge'] = '';
    }

    syncAndRebuildView();
}

// Internal view-refresh utility wrapper
function syncAndRebuildView() {
    draggedAttributeName = null;
    buildStateFromData();
    renderInterface();
}

// Step 5: Implement the Grid Table Component and State Synchronization
// --- Global Search and Filter Variables ---
let columnFilters = {
    'Parent Attribute': '',
    'Attribute': '',
    'Set Status To': '',
    'Recommended Merge': ''
};

// --- Render Table Grid Component ---
function renderTablePreview() {
    const tableEl = document.getElementById('data-table');
    tableEl.innerHTML = '';

    if (globalData.length === 0) {
        tableEl.innerHTML = `<tr><td class="p-4 text-gray-500 italic text-center">No data loaded. Drop an Excel file to begin.</td></tr>`;
        return;
    }

    // Target columns to render
    const headers = ['Parent Attribute', 'Attribute', 'Set Status To', 'Recommended Merge', 'Other Matches'];

    // Assemble the Table Header Row with Filter Input Boxes
    let headerHTML = `<thead class="bg-gray-850 sticky top-0 border-b border-gray-700 z-10"><tr>`;
    headers.forEach(h => {
        headerHTML += `
            <th class="p-2 text-gray-400 font-semibold tracking-wider border-r border-gray-700 min-w-[150px]">
                <div class="flex flex-col space-y-1">
                    <span>${h}</span>
                    ${h !== 'Other Matches' ? `
                    <input type="text" 
                           placeholder="Filter..." 
                           value="${columnFilters[h] || ''}" 
                           oninput="updateColumnFilter('${h}', this.value)" 
                           class="w-full bg-gray-900 border border-gray-700 rounded px-1.5 py-0.5 text-[11px] text-gray-200 focus:outline-none focus:border-indigo-500">` : ''}
                </div>
            </th>`;
    });
    headerHTML += `</tr></thead>`;

    // Render Table Rows matching active filters
    let bodyHTML = `<tbody>`;
    globalData.forEach((row, rowIndex) => {
        // Evaluate filter criteria matches
        if (columnFilters['Parent Attribute'] && !String(row['Parent Attribute'] || '').toLowerCase().includes(columnFilters['Parent Attribute'].toLowerCase())) return;
        if (columnFilters['Attribute'] && !String(row['Attribute'] || '').toLowerCase().includes(columnFilters['Attribute'].toLowerCase())) return;
        if (columnFilters['Set Status To'] && !String(row['Set Status To'] || '').toLowerCase().includes(columnFilters['Set Status To'].toLowerCase())) return;
        if (columnFilters['Recommended Merge'] && !String(row['Recommended Merge'] || '').toLowerCase().includes(columnFilters['Recommended Merge'].toLowerCase())) return;

        bodyHTML += `<tr class="border-b border-gray-800 hover:bg-gray-750/40 transition-colors">`;
        
        headers.forEach(h => {
            const cellValue = row[h] || '';
            
            if (h === 'Other Matches') {
                // Keep the matches column read-only to avoid breaking the JSON format
                bodyHTML += `<td class="p-2 border-r border-gray-800 text-gray-500 truncate max-w-xs" title="${cellValue}">${cellValue}</td>`;
            } else if (h === 'Set Status To') {
                // Dropdown selector for status column validation
                bodyHTML += `
                    <td class="p-1 border-r border-gray-800">
                        <select onchange="updateTableCell(${rowIndex}, '${h}', this.value)" class="w-full bg-gray-900 border border-gray-700 rounded px-1 py-0.5 text-xs text-gray-200 focus:outline-none focus:border-indigo-500">
                            <option value="MERGE" ${cellValue === 'MERGE' ? 'selected' : ''}>MERGE</option>
                            <option value="NAV" ${cellValue === 'NAV' ? 'selected' : ''}>NAV</option>
                            <option value="KEEP" ${cellValue === 'KEEP' ? 'selected' : ''}>KEEP</option>
                            <option value="DROP" ${cellValue === 'DROP' ? 'selected' : ''}>DROP</option>
                            <option value="WIP" ${cellValue === 'WIP' ? 'selected' : ''}>WIP</option>
                        </select>
                    </td>`;
            } else {
                // Editable input text cells
                bodyHTML += `
                    <td class="p-1 border-r border-gray-800">
                        <input type="text" 
                               value="${cellValue}" 
                               onchange="updateTableCell(${rowIndex}, '${h}', this.value)" 
                               class="w-full bg-transparent px-2 py-1 text-xs text-gray-200 focus:bg-gray-900 focus:outline-none focus:ring-1 focus:ring-indigo-500 rounded">
                    </td>`;
            }
        });
        bodyHTML += `</tr>`;
    });
    bodyHTML += `</tbody>`;

    tableEl.innerHTML = headerHTML + bodyHTML;
}

// --- Live Input Change Listeners ---
function updateTableCell(rowIndex, columnName, value) {
    globalData[rowIndex][columnName] = value;
}

function updateColumnFilter(columnName, value) {
    columnFilters[columnName] = value;
    renderTablePreview(); // Re-filter the rows without recreating the trees
}

// --- Synchronizing Event Triggers ---

// Button: Refresh Interface (Table -> Visual Layouts)
document.getElementById('btn-refresh').addEventListener('click', () => {
    buildStateFromData();
    renderInterface();
});

// Button: Save State (Visual Layouts -> Table Content)
document.getElementById('btn-save').addEventListener('click', () => {
    // Sync table contents with changes from drag-and-drop
    globalData.forEach(row => {
        const currentAttr = row['Attribute'];
        
        // Check if the item is inside a tree structure
        let foundInTree = false;
        Object.keys(treeStructures).forEach(rootName => {
            if (treeStructures[rootName].includes(currentAttr)) {
                row['Set Status To'] = 'MERGE';
                row['Recommended Merge'] = rootName;
                foundInTree = true;
            }
        });

        // Check if the item is a root element
        if (Object.keys(treeStructures).includes(currentAttr)) {
            row['Recommended Merge'] = '';
            // If it has children, keep its current status; otherwise, set it to WIP
            if (treeStructures[currentAttr].length === 0 && row['Set Status To'] === 'MERGE') {
                row['Set Status To'] = 'WIP';
            }
            foundInTree = true;
        }

        // If the item is in the ungrouped list, clear its parent reference
        if (!foundInTree) {
            const ungroupedMatch = ungroupedList.find(item => item.name === currentAttr);
            if (ungroupedMatch) {
                row['Set Status To'] = ungroupedMatch.status;
                row['Recommended Merge'] = '';
            }
        }
    });

    renderTablePreview();
    alert('Application state successfully synchronized with the data table.');
});

// Step 6: Exporting the Edited Data Back to Excel
// --- Download Excel Workbook ---
document.getElementById('btn-download').addEventListener('click', () => {
    if (globalData.length === 0) {
        alert("There is no data available to download.");
        return;
    }

    // Final clean pass to ensure parent attributes match the specifications
    globalData.forEach(row => {
        const attrName = row['Attribute'];
        
        // If it is an active root node with children, map its requirements
        if (treeStructures[attrName] && treeStructures[attrName].length > 0) {
            treeStructures[attrName].forEach(childName => {
                let childRow = globalData.find(r => r['Attribute'] === childName);
                if (childRow) {
                    childRow['Parent Attribute'] = attrName;
                }
            });
        }
    });

    // Convert the data back to an Excel sheet layout
    const worksheet = XLSX.utils.json_to_sheet(globalData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Deduplicated Attributes");

    // Save and download the file
    XLSX.writeFile(workbook, "Deduplicated_Attributes_Export.xlsx");
});

