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
    
    // File change handler via standard browsing window click
    fileInput.addEventListener("change", (e) => {
        if (e.target.files && e.target.files.length > 0) {
            handleFile(e.target.files[0]); // FIXED: Isolate the raw file object directly
        }
    });

    // Prevent browser from opening/downloading the file natively on ALL drag phases
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
        }, false);
    });

    // Provide immediate visual confirmation hover highlights
    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => {
            dropzone.style.background = "#e0e7ff";
            dropzone.style.borderColor = "#4f46e5";
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, () => {
            dropzone.style.background = "#ffffff";
            dropzone.style.borderColor = "#a5b4fc";
        }, false);
    });

    // Handle dropped files
    dropzone.addEventListener("drop", (e) => {
        const dt = e.dataTransfer;
        if (dt.files && dt.files.length > 0) {
            handleFile(dt.files[0]); // FIXED: Isolate the raw file object directly
        }
    }, false);

    // Panel collapsing toggles
    document.getElementById("toggle-ungrouped").addEventListener("click", () => {
        document.getElementById("ungrouped-panel").classList.toggle("collapsed");
    });
    document.getElementById("toggle-table").addEventListener("click", () => {
        document.getElementById("table-panel").classList.toggle("uncollapsed");
        if (tabulatorTable) tabulatorTable.redraw(); 
    });

    // Action button triggers
    document.getElementById("refresh-btn").addEventListener("click", () => rebuildUIFromMemory());
    document.getElementById("save-btn").addEventListener("click", () => syncTreesToMemory());
    document.getElementById("download-btn").addEventListener("click", () => exportToExcel());
}

// 1. DATA INGESTION: Read uploaded Excel sheet into memory
function handleFile(fileObject) {
    if (!fileObject) return;
    
    console.log("Running FileReader on object:", fileObject.name);
    const reader = new FileReader();
    
    reader.onload = (e) => {
        try {
            // Confirm library initialization state flags directly before processing parsing pipeline
            if (typeof XLSX === 'undefined') {
                throw new Error("XLSX core engine library failed to load due to network security rules.");
            }

            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // FIXED: Isolates first index string directly
            const firstSheetName = workbook.SheetNames[0]; 
            console.log("Accessing spreadsheet sheet page title:", firstSheetName);
            
            const worksheet = workbook.Sheets[firstSheetName];
            if (!worksheet) {
                throw new Error("Target sheet dataset structure evaluates as empty.");
            }
            
            // Convert rows into a raw JSON array of objects
            globalData = XLSX.utils.sheet_to_json(worksheet);
            console.log("Successfully extracted rows into memory:", globalData.length);
            
            if (globalData.length === 0) {
                alert("The uploaded Excel sheet contains no valid row entries.");
                return;
            }

            // Transition interfaces screens
            document.getElementById("upload-screen").classList.remove("active");
            document.getElementById("dashboard-screen").classList.add("active");
            
            // Populate spreadsheet views
            buildTabulatorTable();
            rebuildUIFromMemory();
        } catch (error) {
            console.error("Critical Runtime Ingestion Exception:", error);
            alert("File processing failed: " + error.message);
        }
    };
    reader.readAsArrayBuffer(fileObject);
}

// 2. THE SPREADSHEET ENGINE: Initialize Tabulator configuration
function buildTabulatorTable() {
    if (typeof Tabulator === 'undefined') return;
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

    const mergeRows = globalData.filter(row => row["Set Status To"] === "MERGE");
    const ungroupedRows = globalData.filter(row => row["Set Status To"] !== "MERGE");

    // Populate Right Sidebar (Ungrouped Items)
    ungroupedRows.forEach(row => {
        if (!row["Attribute"]) return;
        const item = document.createElement("div");
        item.className = `ungrouped-item ${row["Set Status To"] || 'WIP'}`;
        item.innerText = row["Attribute"];
        item.setAttribute("draggable", "true");
        
        item.addEventListener("dragstart", (e) => {
            e.dataTransfer.setData("text/plain", JSON.stringify({ source: 'UNGROUPED', attribute: row["Attribute"] }));
        });
        
        ungroupedList.appendChild(item);
    });

    // Group matching leaves cleanly under their unique "Recommended Merge" target root node
    const treeMap = {};
    mergeRows.forEach(row => {
        const root = row["Recommended Merge"] || "Unspecified Target";
        if (!treeMap[root]) treeMap[root] = [];
        treeMap[root].push(row);
    });

    // Generate functional HTML containers for each individual relationship group
    let treeCounter = 0;
    if (typeof Treant !== 'undefined') {
        for (const [rootNodeName, children] of Object.entries(treeMap)) {
            const treeContainerId = `tree-canvas-${treeCounter++}`;
            
            const card = document.createElement("div");
            card.className = "tree-card";
            card.innerHTML = `<div id="${treeContainerId}" class="tree-canvas-render"></div>`;
            treeGrid.appendChild(card);

            const treantConfig = {
                chart: { container: `#${treeContainerId}`, connecters: { type: "step" } },
                nodeStructure: {
                    text: { name: rootNodeName },
                    HTMLclass: "root-node-style",
                    children: children.map(child => ({
                        text: { name: child["Attribute"] },
                        HTMLclass: "leaf-node-style",
                        HTMLid: `node-${btoa(encodeURIComponent(child["Attribute"]))}`, 
                        dataAttributes: { attr: child["Attribute"] }
                    }))
                }
            };

            new Treant(treantConfig);
            attachNodeInteractivity(children);
        }
    }
}

// 4. INTERACTIVITY LOGIC: Alternate selection overlay handler
function attachNodeInteractivity(children) {
    children.forEach(child => {
        const encodedId = `node-${btoa(encodeURIComponent(child["Attribute"]))}`;
        const nodeElement = document.getElementById(encodedId);
        if (!nodeElement) return;

        nodeElement.addEventListener("mouseenter", (e) => {
            removeExistingTooltips();
            
            const alternates = parseOtherMatches(child["Other Matches"]);
            if (alternates.length === 0) return;

            const tooltip = document.createElement("div");
            tooltip.className = "node-alternates-tooltip";
            tooltip.innerHTML = `<strong>Change Parent to:</strong><hr>`;
            
            alternates.forEach(alt => {
                const link = document.createElement("div");
                link.className = "tooltip-alt-link";
                link.innerText = `${alt.name} (${Math.round(alt.score * 100)}%)`;
                
                link.addEventListener("click", () => {
                    updateAttributeParent(child["Attribute"], alt.name);
                    removeExistingTooltips();
                });
                tooltip.appendChild(link);
            });

            document.body.appendChild(tooltip);
            const rect = nodeElement.getBoundingClientRect();
            tooltip.style.left = `${rect.left + window.scrollX}px`;
            tooltip.style.top = ${rect.bottom + window.scrollY + 5}px;
        });
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
function parseOtherMatches(rawString) {
    if (!rawString) return [];
    try {
        let clean = rawString.trim();
        if(clean.startsWith("{") && clean.endsWith("}")) {
            const cleanJson = JSON.parse(clean);
            return Object.entries(cleanJson).map(([name, score]) => ({ name, score }));
        }
        return [];
    } 
    catch(err) {
        console.warn("Could not extract alternate dataset configurations", err);
        return [];
    }
}
function removeExistingTooltips() {
    document.querySelectorAll(".node-alternates-tooltip").forEach(t => t.remove());
}
// 5. DATA STATE OPERATIONS: Re-assign parent referencesfunction update
AttributeParent(attributeName, newParentName) {
    const targetRow = globalData.find(row => row["Attribute"] === attributeName);
    if (targetRow) {
        targetRow["Recommended Merge"] = newParentName;
        targetRow["Set Status To"] = "MERGE";
        if (tabulatorTable) tabulatorTable.setData(globalData);
        rebuildUIFromMemory();
    }
}
// 6. SYNCHRONIZATION AND WRITING: Collect visual changes back into the final pipeline data grid
function syncTreesToMemory() {
    if (tabulatorTable) {
        globalData = tabulatorTable.getData();
        rebuildUIFromMemory();
        alert("State variables synchronized successfully.");
    }
}
// 7. FILE EXPORT PIPELINE: Write active data matrices straight to dynamic spreadsheet files
function exportToExcel() {
    const worksheet = XLSX.utils.json_to_sheet(globalData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Updated Attributes");
    XLSX.writeFile(workbook, "Attribute_Deduplication_Output.xlsx");
}
