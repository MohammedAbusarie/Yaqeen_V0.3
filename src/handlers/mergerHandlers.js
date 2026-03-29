import {
  ValidationError,
  DownloadError,
  FileError,
  fetchXlsxFromUrl,
  readWorkbookFromArrayBuffer,
} from "../../attendance.js";
import {
  extractAllColumns,
  buildMappingMatrix,
  mergeColumnsSequentially,
  generateMergedWorkbook,
} from "../sheetMerger.js";
import { readFileAsArrayBuffer } from "../fileRead.js";
import { safeBaseName } from "../metadata.js";
import { downloadBlob } from "../dom.js";

export function createMergerHandlers(refs) {
  const { els, state } = refs;

  // ============================================================================
  // Sheet Merger Handlers
  // ============================================================================

  /**
   * Set Sheet Merger status message
   */
  function setMergerStatus(msg, kind = "info") {
    const statusEl = els.mergerStatus;
    if (!statusEl) return;
    
    statusEl.textContent = msg || "";
    statusEl.classList.remove("is-error", "is-ok");
    if (kind === "error") statusEl.classList.add("is-error");
    if (kind === "ok") statusEl.classList.add("is-ok");
  }

  /**
   * Load workbook for sheet merger
   */
  async function handleMergerLoadFile() {
    try {
      setMergerStatus("Loading workbook...");
      
      // Try to load from file upload first, then from URL
      let arrayBuffer = null;
      let fileName = null;

      const file = els.mergerXlsxFile?.files?.[0];
      if (file) {
        try {
          arrayBuffer = await readFileAsArrayBuffer(file);
          fileName = file.name;
        } catch (err) {
          if (file.name.endsWith(".ods")) {
            throw new FileError("ODS file upload failed. Please convert the file to XLSX format and upload again.");
          }
          throw new FileError(`Failed to read file: ${err.message}`);
        }
      } else {
        // Try URL download
        const url = String(els.mergerSheetUrl?.value || "").trim();
        if (!url) {
          throw new ValidationError("Please upload a file or provide a Google Sheet URL.");
        }

        try {
          arrayBuffer = await fetchXlsxFromUrl(url);
          fileName = "downloaded-sheet.xlsx";
        } catch (err) {
          if (err instanceof DownloadError || err.message.includes("CORS")) {
            throw new DownloadError("Could not download from Google Sheets. Please download the file manually: File → Download → Microsoft Excel (.xlsx), then upload it above.");
          }
          throw err;
        }
      }

      // Parse workbook
      const workbook = readWorkbookFromArrayBuffer(arrayBuffer);
      if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
        throw new ValidationError("Workbook contains no sheets.");
      }

      // Extract all columns
      const allColumns = extractAllColumns(workbook);
      if (allColumns.length === 0) {
        throw new ValidationError("No columns found in workbook.");
      }

      // Update state
      state.sheetMerger.workbookArrayBuffer = arrayBuffer;
      state.sheetMerger.workbookName = fileName;
      state.sheetMerger.workbookLoaded = true;
      state.sheetMerger.workbookSheetNames = workbook.SheetNames;
      state.sheetMerger.allColumns = allColumns;
      // Reset maxPositions to 10 when loading a new workbook
      state.sheetMerger.maxPositions = 10;
      state.sheetMerger.mapping = buildMappingMatrix(workbook.SheetNames, state.sheetMerger.maxPositions);
      // Generate colors for sheets (reset when loading new workbook)
      state.sheetMerger.sheetColors = {};

      // Render column pool and mapping matrix
      renderColumnPoolGrouped();
      renderMappingMatrix();

      // Show mapping step
      if (els.mergerMappingStep) {
        els.mergerMappingStep.style.display = "block";
      }

      setMergerStatus(`Loaded ${workbook.SheetNames.length} sheet(s) with ${allColumns.length} column(s). Drag columns to the matrix to start mapping.`, "ok");
    } catch (err) {
      console.error("Sheet merger load error:", err);
      setMergerStatus(err.message || "Failed to load workbook", "error");
    }
  }

  /**
   * Generate a distinct color for a sheet based on its index
   * @param {number} index - 0-based index of the sheet
   * @returns {string} CSS color string (HSL)
   */
  function getSheetColor(index) {
    // Generate distinct colors using HSL
    // Use golden angle approximation for better color distribution
    const goldenAngle = 137.508;
    const hue = (index * goldenAngle) % 360;
    // Use good saturation and lightness for visibility
    const saturation = 65;
    const lightness = 55;
    return `hsl(${hue}, ${saturation}%, ${lightness}%)`;
  }

  /**
   * Get or generate colors for all sheets
   * @returns {Record<string, string>} Map of sheet name to color
   */
  function getSheetColors() {
    const sheets = state.sheetMerger.workbookSheetNames || [];
    const existingColors = state.sheetMerger.sheetColors || {};
    
    // Check if we need to generate colors (empty object or missing sheets)
    const needsGeneration = Object.keys(existingColors).length === 0 || 
                           sheets.some(sheetName => !existingColors[sheetName]);
    
    if (needsGeneration) {
      state.sheetMerger.sheetColors = {};
      sheets.forEach((sheetName, index) => {
        state.sheetMerger.sheetColors[sheetName] = getSheetColor(index);
      });
    }
    
    return state.sheetMerger.sheetColors || {};
  }

  /**
   * Render column pool with accordion grouping by sheet
   */
  function renderColumnPoolGrouped() {
    const poolEl = els.mergerColumnPool;
    if (!poolEl) return;

    poolEl.innerHTML = "";

    const columns = state.sheetMerger.allColumns || [];
    const sheets = state.sheetMerger.workbookSheetNames || [];
    const expandedSheets = state.sheetMerger.expandedSheets || [sheets[0]]; // First sheet expanded by default
    const searchQuery = (state.sheetMerger.searchQuery || "").toLowerCase();
    const sheetColors = getSheetColors();

    // Group columns by sheet
    const columnsBySheet = {};
    for (const col of columns) {
      if (!columnsBySheet[col.sheet]) {
        columnsBySheet[col.sheet] = [];
      }
      columnsBySheet[col.sheet].push(col);
    }

    // Render each sheet group
    for (const sheetName of sheets) {
      const sheetColumns = columnsBySheet[sheetName] || [];
      
      // Filter columns based on search
      const filteredColumns = searchQuery 
        ? sheetColumns.filter(col => 
            col.headerText.toLowerCase().includes(searchQuery) ||
            col.sampleValues.some(val => val.toLowerCase().includes(searchQuery))
          )
        : sheetColumns;

      // Skip sheet if no matching columns when searching
      if (searchQuery && filteredColumns.length === 0) continue;

      const isExpanded = expandedSheets.includes(sheetName) || searchQuery;
      const sheetColor = sheetColors[sheetName] || getSheetColor(0);
      
      // Create group container
      const group = document.createElement("div");
      group.className = `columnGroup ${isExpanded ? 'columnGroup--expanded' : 'columnGroup--collapsed'}`;
      
      // Create header
      const header = document.createElement("div");
      header.className = "columnGroup__header";
      header.dataset.sheet = sheetName;
      // Apply sheet color as left border
      header.style.borderLeftWidth = "4px";
      header.style.borderLeftStyle = "solid";
      header.style.borderLeftColor = sheetColor;
      
      const countText = searchQuery && filteredColumns.length !== sheetColumns.length
        ? `(${filteredColumns.length} of ${sheetColumns.length})`
        : `(${sheetColumns.length})`;
      
      header.innerHTML = `
        <span class="columnGroup__icon">${isExpanded ? '▼' : '▶'}</span>
        <span class="columnGroup__name">${sheetName}</span>
        <span class="columnGroup__count">${countText}</span>
      `;
      
      // Create body
      const body = document.createElement("div");
      body.className = "columnGroup__body";
      if (!isExpanded) body.style.display = "none";
      
      // Add columns to body
      for (const col of filteredColumns) {
        const item = document.createElement("div");
        item.className = "columnItem";
        item.draggable = true;
        item.dataset.columnKey = col.key;
        // Apply sheet color as left border accent
        item.style.borderLeftWidth = "3px";
        item.style.borderLeftStyle = "solid";
        item.style.borderLeftColor = sheetColor;
        
        const sampleText = col.sampleValues.length > 0 
          ? `<div class="columnItem__sample">${col.sampleValues.slice(0, 3).join(", ")}...</div>`
          : "";

        item.innerHTML = `
          <div class="columnItem__header">
            <strong>${col.headerText}</strong>
            <span class="columnItem__meta">${col.columnLetter}</span>
          </div>
          ${sampleText}
        `;

        body.appendChild(item);
      }
      
      group.appendChild(header);
      group.appendChild(body);
      poolEl.appendChild(group);
    }

    // Show "no results" message if searching and nothing found
    if (searchQuery && poolEl.children.length === 0) {
      const noResults = document.createElement("div");
      noResults.className = "columnPool__noResults";
      noResults.textContent = `No columns found matching "${state.sheetMerger.searchQuery}"`;
      poolEl.appendChild(noResults);
    }
  }

  /**
   * Handle column search input
   */
  function handleMergerColumnSearch(e) {
    state.sheetMerger.searchQuery = e.target.value;
    renderColumnPoolGrouped();
  }

  /**
   * Handle column group toggle (expand/collapse)
   */
  function handleMergerColumnGroupToggle(e) {
    const header = e.target.closest(".columnGroup__header");
    if (!header) return;
    
    const sheetName = header.dataset.sheet;
    const group = header.closest(".columnGroup");
    const body = group.querySelector(".columnGroup__body");
    const icon = group.querySelector(".columnGroup__icon");
    
    if (!state.sheetMerger.expandedSheets) {
      state.sheetMerger.expandedSheets = [];
    }
    
    if (state.sheetMerger.expandedSheets.includes(sheetName)) {
      // Collapse
      state.sheetMerger.expandedSheets = state.sheetMerger.expandedSheets.filter(s => s !== sheetName);
      group.classList.remove("columnGroup--expanded");
      group.classList.add("columnGroup--collapsed");
      body.style.display = "none";
      icon.textContent = "▶";
    } else {
      // Expand
      state.sheetMerger.expandedSheets.push(sheetName);
      group.classList.add("columnGroup--expanded");
      group.classList.remove("columnGroup--collapsed");
      body.style.display = "block";
      icon.textContent = "▼";
    }
  }

  /**
   * Render mapping matrix with drop zones
   */
  function renderMappingMatrix() {
    const matrixEl = els.mergerMappingMatrix;
    if (!matrixEl) return;

    matrixEl.innerHTML = "";

    const sheetNames = state.sheetMerger.workbookSheetNames || [];
    const mapping = state.sheetMerger.mapping || {};
    const maxPositions = state.sheetMerger.maxPositions || 10;

    // Update column control buttons
    const controlsEl = els.mergerColumnControls;
    if (controlsEl) {
      controlsEl.innerHTML = "";
      
      const addBtn = document.createElement("button");
      addBtn.type = "button";
      addBtn.className = "btn";
      addBtn.style.fontSize = "13px";
      addBtn.style.padding = "8px 14px";
      addBtn.textContent = "+ Add Column";
      addBtn.title = "Add a new column at the end";
      addBtn.addEventListener("click", handleMergerAddColumn);
      controlsEl.appendChild(addBtn);
      
      if (maxPositions > 1) {
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "btn btn--ghost";
        removeBtn.style.fontSize = "13px";
        removeBtn.style.padding = "8px 14px";
        removeBtn.textContent = "− Remove Column";
        removeBtn.title = "Remove the last column (clears mappings in last position)";
        removeBtn.addEventListener("click", handleMergerRemoveColumn);
        controlsEl.appendChild(removeBtn);
      }
    }

    // Build header row
    const headerRow = document.createElement("div");
    headerRow.className = "mappingMatrix__row mappingMatrix__row--header";
    // Set dynamic grid columns: 120px for label + maxPositions columns of 100px each
    headerRow.style.gridTemplateColumns = `120px repeat(${maxPositions}, 100px)`;
    
    const cornerCell = document.createElement("div");
    cornerCell.className = "mappingMatrix__cell mappingMatrix__cell--corner";
    cornerCell.textContent = "Sheet \\ Position";
    headerRow.appendChild(cornerCell);

    for (let pos = 0; pos < maxPositions; pos++) {
      const headerCell = document.createElement("div");
      headerCell.className = "mappingMatrix__cell mappingMatrix__cell--header";
      headerCell.textContent = `Col ${pos + 1}`;
      headerRow.appendChild(headerCell);
    }
    
    matrixEl.appendChild(headerRow);

    // Build data rows
    for (const sheetName of sheetNames) {
      const row = document.createElement("div");
      row.className = "mappingMatrix__row";
      // Set dynamic grid columns: 120px for label + maxPositions columns of 100px each
      row.style.gridTemplateColumns = `120px repeat(${maxPositions}, 100px)`;
      
      const labelCell = document.createElement("div");
      labelCell.className = "mappingMatrix__cell mappingMatrix__cell--label";
      labelCell.textContent = sheetName;
      row.appendChild(labelCell);

      for (let pos = 0; pos < maxPositions; pos++) {
        const dropZone = document.createElement("div");
        dropZone.className = "mappingMatrix__cell mappingMatrix__cell--dropzone";
        dropZone.dataset.sheet = sheetName;
        dropZone.dataset.position = String(pos);

        const columnKey = mapping[sheetName]?.[pos];
        if (columnKey) {
          const col = state.sheetMerger.allColumns.find(c => c.key === columnKey);
          if (col) {
            dropZone.classList.add("mappingMatrix__cell--filled");
            const sheetColors = getSheetColors();
            const sheetColor = sheetColors[col.sheet] || getSheetColor(0);
            // Apply sheet color as background with good contrast
            dropZone.style.backgroundColor = sheetColor;
            dropZone.style.opacity = "0.9";
            dropZone.style.borderColor = sheetColor;
            dropZone.style.borderWidth = "2px";
            dropZone.innerHTML = `
              <div class="mappedColumn" style="color: white; text-shadow: 0 1px 2px rgba(0,0,0,0.3);">
                <div class="mappedColumn__header" style="color: white;">${col.headerText}</div>
                <div class="mappedColumn__meta" style="color: rgba(255,255,255,0.9);">${col.columnLetter}</div>
                <button class="mappedColumn__remove" data-sheet="${sheetName}" data-position="${pos}" type="button">×</button>
              </div>
            `;
          }
        } else {
          dropZone.textContent = "Drop here";
        }

        row.appendChild(dropZone);
      }
      matrixEl.appendChild(row);
    }
  }

  /**
   * Handle column drag start
   */
  function handleMergerColumnDragStart(e) {
    const target = e.target.closest(".columnItem");
    if (!target) return;

    const columnKey = target.dataset.columnKey;
    if (!columnKey) return;

    e.dataTransfer.effectAllowed = "copy";
    e.dataTransfer.setData("text/plain", columnKey);
    target.classList.add("columnItem--dragging");

    // Store in state for fallback
    state.sheetMerger.draggedColumnKey = columnKey;
  }

  /**
   * Handle drag over matrix cell
   */
  function handleMergerMatrixDragOver(e) {
    const target = e.target.closest(".mappingMatrix__cell--dropzone");
    if (!target) return;

    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
    target.classList.add("mappingMatrix__cell--dragover");
  }

  /**
   * Handle drop on matrix cell
   */
  function handleMergerMatrixDrop(e) {
    const target = e.target.closest(".mappingMatrix__cell--dropzone");
    if (!target) return;

    e.preventDefault();
    target.classList.remove("mappingMatrix__cell--dragover");

    const sheet = target.dataset.sheet;
    const position = parseInt(target.dataset.position, 10);
    
    // Get column key from drag data or fallback to state
    let columnKey = e.dataTransfer.getData("text/plain");
    if (!columnKey) {
      columnKey = state.sheetMerger.draggedColumnKey;
    }

    if (!columnKey || !sheet || isNaN(position)) return;

    // Update mapping
    if (!state.sheetMerger.mapping[sheet]) {
      state.sheetMerger.mapping[sheet] = {};
    }
    state.sheetMerger.mapping[sheet][position] = columnKey;

    // Re-render matrix
    renderMappingMatrix();

    // Clear dragging state
    delete state.sheetMerger.draggedColumnKey;
    const draggingItems = document.querySelectorAll(".columnItem--dragging");
    draggingItems.forEach(item => item.classList.remove("columnItem--dragging"));
  }

  /**
   * Handle click on matrix (for remove buttons)
   */
  function handleMergerMatrixClick(e) {
    const removeBtn = e.target.closest(".mappedColumn__remove");
    if (!removeBtn) return;

    const sheet = removeBtn.dataset.sheet;
    const position = parseInt(removeBtn.dataset.position, 10);

    if (sheet && !isNaN(position) && state.sheetMerger.mapping[sheet]) {
      state.sheetMerger.mapping[sheet][position] = null;
      renderMappingMatrix();
    }
  }

  /**
   * Handle add column button click
   */
  function handleMergerAddColumn() {
    if (!state.sheetMerger.maxPositions) {
      state.sheetMerger.maxPositions = 10;
    }
    state.sheetMerger.maxPositions++;
    
    // Ensure mapping structure exists for all sheets
    const sheetNames = state.sheetMerger.workbookSheetNames || [];
    for (const sheetName of sheetNames) {
      if (!state.sheetMerger.mapping[sheetName]) {
        state.sheetMerger.mapping[sheetName] = {};
      }
      // Initialize new position to null if it doesn't exist
      const lastPos = state.sheetMerger.maxPositions - 1;
      if (state.sheetMerger.mapping[sheetName][lastPos] === undefined) {
        state.sheetMerger.mapping[sheetName][lastPos] = null;
      }
    }
    
    renderMappingMatrix();
  }

  /**
   * Handle remove column button click
   */
  function handleMergerRemoveColumn() {
    if (!state.sheetMerger.maxPositions || state.sheetMerger.maxPositions <= 1) {
      return; // Don't allow removing the last column
    }
    
    const lastPos = state.sheetMerger.maxPositions - 1;
    
    // Clear any mappings in the last position across all sheets
    const sheetNames = state.sheetMerger.workbookSheetNames || [];
    for (const sheetName of sheetNames) {
      if (state.sheetMerger.mapping[sheetName] && state.sheetMerger.mapping[sheetName][lastPos] !== undefined) {
        delete state.sheetMerger.mapping[sheetName][lastPos];
      }
    }
    
    state.sheetMerger.maxPositions--;
    renderMappingMatrix();
  }

  /**
   * Handle eliminate headers checkbox
   */
  function handleMergerEliminateHeaders(e) {
    state.sheetMerger.eliminateHeaders = e.target.checked;
  }

  /**
   * Generate preview of merged data
   */
  async function handleMergerGeneratePreview() {
    try {
      setMergerStatus("Generating preview...");

      if (!state.sheetMerger.workbookLoaded || !state.sheetMerger.workbookArrayBuffer) {
        throw new ValidationError("Please load a workbook first.");
      }

      // Check if any columns are mapped
      let hasMapping = false;
      const mapping = state.sheetMerger.mapping || {};
      for (const sheet in mapping) {
        for (const pos in mapping[sheet]) {
          if (mapping[sheet][pos]) {
            hasMapping = true;
            break;
          }
        }
        if (hasMapping) break;
      }

      if (!hasMapping) {
        throw new ValidationError("Please map at least one column to a position before generating preview.");
      }

      // Parse workbook
      const workbook = readWorkbookFromArrayBuffer(state.sheetMerger.workbookArrayBuffer);

      // Merge columns
      const mergedData = mergeColumnsSequentially(
        workbook,
        mapping,
        state.sheetMerger.allColumns,
        state.sheetMerger.eliminateHeaders
      );

      // Update state
      state.sheetMerger.mergedData = mergedData;
      state.sheetMerger.previewRows = mergedData.rows;
      // Reset preview rows loaded to initial value
      state.sheetMerger.previewRowsLoaded = 100;

      // Render preview
      renderMergerPreview();

      // Show preview step
      if (els.mergerPreviewStep) {
        els.mergerPreviewStep.style.display = "block";
      }

      // Scroll to preview
      els.mergerPreviewStep?.scrollIntoView({ behavior: "smooth", block: "start" });

      setMergerStatus(`Preview generated successfully. ${mergedData.totalRows} rows × ${mergedData.totalColumns} columns.`, "ok");
    } catch (err) {
      console.error("Merger preview error:", err);
      setMergerStatus(err.message || "Failed to generate preview", "error");
    }
  }

  /**
   * Render preview table
   */
  function renderMergerPreview() {
    const mergedData = state.sheetMerger.mergedData;
    if (!mergedData || !mergedData.rows) return;

    // Update summary
    if (els.mergerPreviewRowCount) els.mergerPreviewRowCount.textContent = String(mergedData.totalRows);
    if (els.mergerPreviewColCount) els.mergerPreviewColCount.textContent = String(mergedData.totalColumns);
    if (els.mergerPreviewSheetCount) els.mergerPreviewSheetCount.textContent = String(mergedData.sourceSheets.length);

    // Render table header
    const thead = els.mergerPreviewTableHead;
    if (thead && mergedData.headers && mergedData.headers.length > 0) {
      thead.innerHTML = "";
      const headerRow = document.createElement("tr");
      mergedData.headers.forEach((header, idx) => {
        const th = document.createElement("th");
        th.textContent = String(header);
        headerRow.appendChild(th);
      });
      thead.appendChild(headerRow);
    }

    // Render table body (limit to previewRowsLoaded rows for performance)
    const tbody = els.mergerPreviewTableBody;
    if (tbody) {
      tbody.innerHTML = "";
      const previewRowsLoaded = state.sheetMerger.previewRowsLoaded || 100;
      const displayRows = mergedData.rows.slice(0, previewRowsLoaded);
      
      displayRows.forEach((row, rowIdx) => {
        // Skip header row if it's the first row
        if (rowIdx === 0 && row.every((val, idx) => val === mergedData.headers[idx])) {
          return;
        }

        const tr = document.createElement("tr");
        row.forEach((cell) => {
          const td = document.createElement("td");
          td.textContent = cell !== null && cell !== undefined ? String(cell) : "";
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });

      if (mergedData.rows.length > previewRowsLoaded) {
        const tr = document.createElement("tr");
        const td = document.createElement("td");
        td.colSpan = mergedData.totalColumns;
        td.style.textAlign = "center";
        td.style.fontStyle = "italic";
        td.style.color = "var(--text-muted)";
        td.textContent = `... and ${mergedData.rows.length - previewRowsLoaded} more rows (showing first ${previewRowsLoaded})`;
        tr.appendChild(td);
        tbody.appendChild(tr);
      }
    }
    
    // Update Load More button visibility
    if (els.btnMergerLoadMore) {
      const previewRowsLoaded = state.sheetMerger.previewRowsLoaded || 100;
      const hasMoreRows = mergedData.rows.length > previewRowsLoaded;
      els.btnMergerLoadMore.style.display = hasMoreRows ? "inline-block" : "none";
    }
  }

  /**
   * Load all remaining rows in preview
   */
  function handleMergerLoadMore() {
    const mergedData = state.sheetMerger.mergedData;
    if (!mergedData || !mergedData.rows) return;
    
    // Set to show all rows
    state.sheetMerger.previewRowsLoaded = mergedData.rows.length;
    
    // Re-render preview
    renderMergerPreview();
  }

  /**
   * Go back from preview step to mapping step
   */
  function handleMergerBack() {
    if (els.mergerPreviewStep) {
      els.mergerPreviewStep.style.display = "none";
    }
    if (els.mergerMappingStep) {
      els.mergerMappingStep.style.display = "block";
      els.mergerMappingStep.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  /**
   * Download merged workbook
   */
  async function handleMergerDownload() {
    try {
      setMergerStatus("Generating merged file...");

      const mergedData = state.sheetMerger.mergedData;
      if (!mergedData) {
        throw new ValidationError("Please generate a preview first.");
      }

      const workbook = generateMergedWorkbook(mergedData, "Merged");
      const wbout = window.XLSX.write(workbook, { bookType: "xlsx", type: "array" });

      const baseName = state.sheetMerger.workbookName || "workbook";
      const safeName = safeBaseName(baseName);
      const fileName = `${safeName}_merged.xlsx`;

      downloadBlob(fileName, wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

      setMergerStatus("Merged file downloaded successfully!", "ok");
    } catch (err) {
      console.error("Merger download error:", err);
      setMergerStatus(err.message || "Failed to download merged file", "error");
    }
  }

  /**
   * Auto-fill remaining sheet rows based on first sheet's mappings
   * Matches columns by name (case-insensitive) and position (±1 column)
   */
  function handleMergerAutoFill() {
    const sheetNames = state.sheetMerger.workbookSheetNames || [];
    const allColumns = state.sheetMerger.allColumns || [];
    const mapping = state.sheetMerger.mapping || {};

    if (sheetNames.length < 2) {
      setMergerStatus("Auto-fill requires at least 2 sheets in the workbook.", "error");
      return;
    }

    const firstSheetName = sheetNames[0];
    const firstSheetMapping = mapping[firstSheetName] || {};

    // Check if first sheet has any mappings
    const firstSheetHasMappings = Object.keys(firstSheetMapping).some(
      pos => firstSheetMapping[pos] !== null && firstSheetMapping[pos] !== undefined
    );

    if (!firstSheetHasMappings) {
      setMergerStatus("Please map at least one column in the first sheet before using auto-fill.", "error");
      return;
    }

    let filledCount = 0;

    // Build a lookup map for columns by sheet and key
    const columnLookup = {};
    for (const col of allColumns) {
      if (!columnLookup[col.sheet]) {
        columnLookup[col.sheet] = {};
      }
      columnLookup[col.sheet][col.key] = col;
    }

    // Iterate through each mapped position in the first sheet
    for (const posStr in firstSheetMapping) {
      const position = parseInt(posStr, 10);
      const columnKey = firstSheetMapping[posStr];

      if (!columnKey) continue; // Skip empty positions

      const firstSheetColumn = columnLookup[firstSheetName]?.[columnKey];
      if (!firstSheetColumn) continue;

      const headerText = firstSheetColumn.headerText.toLowerCase();
      const columnIndex = firstSheetColumn.columnIndex;

      // For each remaining sheet, try to find a matching column
      for (let i = 1; i < sheetNames.length; i++) {
        const currentSheetName = sheetNames[i];

        // Skip if position is already filled
        if (mapping[currentSheetName] && mapping[currentSheetName][position]) {
          continue;
        }

        // Find matching column in current sheet
        // Priority: name match first, then check position within ±1
        const currentSheetColumns = allColumns.filter(col => col.sheet === currentSheetName);
        let matchingColumn = null;

        for (const col of currentSheetColumns) {
          // Check name match (case-insensitive)
          if (col.headerText.toLowerCase() === headerText) {
            // Check position match (within ±1 column)
            if (Math.abs(col.columnIndex - columnIndex) <= 1) {
              matchingColumn = col;
              break; // Found exact match, stop searching
            }
          }
        }

        // If we found a match and the position is empty, map it
        if (matchingColumn) {
          if (!mapping[currentSheetName]) {
            mapping[currentSheetName] = {};
          }
          mapping[currentSheetName][position] = matchingColumn.key;
          filledCount++;
        }
      }
    }

    // Re-render matrix to show the changes
    renderMappingMatrix();

    if (filledCount > 0) {
      setMergerStatus(`Auto-fill completed. Mapped ${filledCount} column(s) across remaining sheets.`, "ok");
    } else {
      setMergerStatus("Auto-fill completed. No matching columns found in remaining sheets.", "info");
    }
  }

  /**
   * Reset sheet merger to initial state
   */
  function handleMergerReset() {
    // Reset state
    state.sheetMerger = {
      workbookArrayBuffer: null,
      workbookName: null,
      workbookLoaded: false,
      workbookSheetNames: [],
      allColumns: [],
      mapping: {},
      mergedData: null,
      previewRows: [],
      eliminateHeaders: false,
      expandedSheets: [],
      searchQuery: "",
      maxPositions: 10,
    };

    // Reset UI
    if (els.mergerSheetUrl) els.mergerSheetUrl.value = "";
    if (els.mergerXlsxFile) els.mergerXlsxFile.value = "";
    if (els.mergerEliminateHeaders) els.mergerEliminateHeaders.checked = false;
    if (els.mergerColumnPool) els.mergerColumnPool.innerHTML = "";
    if (els.mergerMappingMatrix) els.mergerMappingMatrix.innerHTML = "";
    if (els.mergerPreviewTableHead) els.mergerPreviewTableHead.innerHTML = "";
    if (els.mergerPreviewTableBody) els.mergerPreviewTableBody.innerHTML = "";
    if (els.mergerMappingStep) els.mergerMappingStep.style.display = "none";
    if (els.mergerPreviewStep) els.mergerPreviewStep.style.display = "none";

    setMergerStatus("Reset complete. Upload a new workbook to start over.", "ok");
  }

  return {
    handleMergerLoadFile,
    handleMergerColumnDragStart,
    handleMergerMatrixDragOver,
    handleMergerMatrixDrop,
    handleMergerMatrixClick,
    handleMergerEliminateHeaders,
    handleMergerGeneratePreview,
    handleMergerDownload,
    handleMergerReset,
    handleMergerBack,
    handleMergerLoadMore,
    handleMergerColumnSearch,
    handleMergerColumnGroupToggle,
    handleMergerAutoFill,
  };
}
