// Global application state variables
let globalData = [];
let tabulatorTable = null;

document.addEventListener("DOMContentLoaded", () => {
    initEventHandlers();
});

// Initialize basic click, toggle, and dropzone handlers
function initEventHandlers() {
    const dropzone = document.getElementById("dropzone");
    const fileInput = document.getElementById("file-input");

    // Click on dropzone triggers hidden file input
    dropzone.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", (e) => handleFile(e.target.files[0]));

    // --- REPLACED DRAG AND DROP HANDLERS TO PREVENT BROWSER OVERRIDES ---
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
        }, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => {
            dropzone.style.background = "#e0e7ff";
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => {
            dropzone.style.background = "#ffffff";
        }, false);
    });

    dropzone.addEventListener("drop", (e) => {
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    }, false);
    // --- END OF REPLACED DRAG AND DROP HANDLERS ---

    // Panel collapsing toggles
    document.getElementById("toggle-ungrouped").addEventListener("click", () => {
        document.getElementById("ungrouped-panel").classList.toggle("collapsed");
    });
    document.getElementById("toggle-table").addEventListener("click", () => {
        document.getElementById("table-panel").classList.toggle("uncollapsed");
        if (tabulatorTable) tabulatorTable.redraw(); // Force recalculate dimensions on expand
    });

    // Action button triggers
    document.getElementById("refresh-btn").addEventListener("click", () => rebuildUIFromMemory());
    document.getElementById("save-btn").addEventListener("click", () => syncTreesToMemory());
    document.getElementById("download-btn").addEventListener("click", () => exportToExcel());
}

// 1. DATA INGESTION: Read uploaded Excel sheet into memory
function handleFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert rows into a raw JSON array of objects
        globalData = XLSX.utils.sheet_to_json(worksheet);
        
        // Advance view to the dashboard state
        document.getElementById("upload-screen").classList.remove("active");
        document.getElementById("dashboard-screen").classList.add("active");
        
        // Initialize Spreadsheet and visualization layers
        buildTabulatorTable();
        rebuildUIFromMemory();
    };
    reader.readAsArrayBuffer(file);
}

// 2. THE SPREADSHEET ENGINE: Initialize Tabulator configuration
function buildTabulatorTable() {
    tabulatorTable = new Tabulator("#excel-table", {
        data: globalData,
        layout: "fitColumns",
        pagination: "local",
        paginationSize: 10,
        movableColumns: true,
        columns: [
            { title: "Attribute", field: "Attribute", width: 200, headerFilter: "text" },
            { title: "Parent Attribute", field: "Parent", width: 180, editor: "text" },
            { title: "Recommended Merge", field: "Recommended Merge", width: 180, editor: "text" },
            { title: "Set Status To", field: "Set Status To", width: 130, editor: "list", editorParams: { values: ["MERGE", "DROP", "NAV", "KEEP", "WIP"] }, headerFilter: "list", headerFilterParams: { values: ["MERGE", "DROP", "NAV", "KEEP", "WIP"] } },
            { title: "Other Matches", field: "Other Matches", width: 300 }
        ]
    });

    // Keep memory in sync automatically when cells are hand-edited inside the grid
    tabulatorTable.on("cellEdited", function(cell) {
        globalData = tabulatorTable.getData();
    });
}

// 3. GRAPH ENGINE: Construct relationship trees and sidebars from data state
function rebuildUIFromMemory() {
    const treeGrid = document.getElementById("trees-grid");
    const ungroupedList = document.getElementById("ungrouped-list");
    
    treeGrid.innerHTML = "";
    ungroupedList.innerHTML = "";

    // Step A: Separate Merge data from the non-merged metadata entries
    const mergeRows = globalData.filter(row => row["Set Status To"] === "MERGE");
    const ungroupedRows = globalData.filter(row => row["Set Status To"] !== "MERGE");

    // Populate Right Sidebar (Ungrouped Items)
    ungroupedRows.forEach(row => {
        const item = document.createElement("div");
        item.className = `ungrouped-item ${row["Set Status To"] || 'WIP'}`;
        item.innerText = row["Attribute"];
        item.setAttribute("draggable", "true");
        
        // Enable native Drag & Drop source tracking
        item.addEventListener("dragstart", (e) => {
            e.dataTransfer.setData("text/plain", JSON.stringify({ source: 'UNGROUPED', attribute: row["Attribute"] }));
        });
        
        ungroupedList.appendChild(item);
    });

    // Step B: Group matching leaves cleanly under their unique "Recommended Merge" target root node
    const treeMap = {};
    mergeRows.forEach(row => {
        const root = row["Recommended Merge"] || "Unspecified Target";
        if (!treeMap[root]) treeMap[root] = [];
        treeMap[root].push(row);
    });

    // Step C: Generate functional HTML containers for each individual relationship group
    let treeCounter = 0;
    for (const [rootNodeName, children] of Object.entries(treeMap)) {
        const treeContainerId = `tree-canvas-${treeCounter++}`;
        
        const card = document.createElement("div");
        card.className = "tree-card";
        card.innerHTML = `<div id="${treeContainerId}" class="tree-canvas-render"></div>`;
        treeGrid.appendChild(card);

        // Convert the structural loop data to a compatible Treant configuration format
        const treantConfig = {
            chart: { container: `#${treeContainerId}`, connecters: { type: "step" } },
            nodeStructure: {
                text: { name: rootNodeName },
                HTMLclass: "root-node-style",
                children: children.map(child => ({
                    text: { name: child["Attribute"] },
                    HTMLclass: "leaf-node-style",
                    HTMLid: `node-${btoa(encodeURIComponent(child["Attribute"]))}`, // Safe structural tracking handles symbols
                    dataAttributes: { attr: child["Attribute"] }
                }))
            }
        };

        // Render Tree into current card loop frame
        new Treant(treantConfig);
        
        // Attach operational interactivity mappings (Hover Actions for Alternates)
        attachNodeInteractivity(children);
    }
}

// 4. INTERACTIVITY LOGIC: Alternate selection overlay handler
function attachNodeInteractivity(children) {
    children.forEach(child => {
        const encodedId = `node-${btoa(encodeURIComponent(child["Attribute"]))}`;
        const nodeElement = document.getElementById(encodedId);
        if (!nodeElement) return;

        // Mouse Hover: Generate absolute tooltip positioning displaying AI's fallback recommendations
        nodeElement.addEventListener("mouseenter", (e) => {
            removeExistingTooltips();
            
            // Safe JSON string parser parsing alternate lists
            const alternates = parseOtherMatches(child["Other Matches"]);
            if (alternates.length === 0) return;

            const tooltip = document.createElement("div");
            tooltip.className = "node-alternates-tooltip";
            tooltip.innerHTML = `<strong>Change Parent to:</strong><hr>`;
            
            alternates.forEach(alt => {
                const link = document.createElement("div");
                link.className = "tooltip-alt-link";
                link.innerText = `${alt.name} (${Math.round(alt.score * 100)}%)`;
                
                // Swap hierarchy upon choosing fallback match directly
                link.addEventListener("click", () => {
                    updateAttributeParent(child["Attribute"], alt.name);
                    removeExistingTooltips();
                });
                tooltip.appendChild(link);
            });

            document.body.appendChild(tooltip);
            const rect = nodeElement.getBoundingClientRect();
            tooltip.style.left = `${rect.left + window.scrollX}px`;
            tooltip.style.top = `${rect.bottom + window.scrollY + 5}px`;
        });

        // Kill overlay visibility when moving focus completely away
        nodeElement.addEventListener("mouseleave", (e) => {
            setTimeout(() => {
                const activeTooltip = document.querySelector(".node-alternates-tooltip");
                if (activeTooltip && !activeTooltip.matches(':hover')) {
                    removeExistingTooltips();
                }
            }, 300);
        });
    });
}

// Utility: Strip structural brackets out of custom formatting types safely
function parseOtherMatches(rawString) {
    if (!rawString) return [];
    try {
        let clean = rawString.trim();
        if(clean.startsWith("{") && clean.endsWith("}")) {
            const cleanJson = JSON.parse(clean);
            return Object.entries(cleanJson).map(([name, score]) => ({ name, score }));
        }
        return [];
    } catch(err) {
        console.warn("Could not extract alternate dataset configurations", err);
        return [];
    }
}

