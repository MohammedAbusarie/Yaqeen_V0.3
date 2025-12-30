import {
  DownloadError,
  FileError,
  ProcessingError,
  ValidationError,
  applyEditorEdits,
  buildStudentSearchRows,
  fetchXlsxFromGoogleSheetUrl,
  fetchXlsxFromUrl,
  listColumnOptions,
  parseGradesText,
  parseStudentIdsText,
  computeEditorPreview,
  readWorkbookFromArrayBuffer,
  normalizeId,
} from "../attendance.js";
import { safeBaseName } from "./metadata.js";
import { downloadBlob } from "./dom.js";
import { readFileAsArrayBuffer, readFileAsText } from "./fileRead.js";
import { processMultipleImages, generateTextFile } from "./ocr.js";
import {
  extractAllColumns,
  buildMappingMatrix,
  mergeColumnsSequentially,
  generateMergedWorkbook,
  columnIndexToLetter,
} from "./sheetMerger.js";

/**
 * @typedef {import('./state.js').AppState} AppState
 */

/**
 * @param {any} els
 * @param {AppState} state
 * @param {(msg: string, kind?: 'info'|'ok'|'error') => void} setStatus
 * @param {(disabled: boolean, label?: string|null) => void} disableRun
 * @param {(viewName: 'inputs'|'report'|'about'|string) => void} switchView
 */
export function createHandlers({ els, state, setStatus, disableRun, switchView }) {
  let editorActiveFixIndex = null;
  let editorActiveGradeIndex = null;
  /** @type {Array<{sheet:string,row1:number,id:string,name:string}>} */
  let editorSearchRows = [];

  function setEditorStatus(msg, kind = "info") {
    if (!els.editorStatus) return;
    els.editorStatus.textContent = msg || "";
    els.editorStatus.classList.toggle("is-error", kind === "error");
    els.editorStatus.classList.toggle("is-ok", kind === "ok");
  }

  function ensureWorkbookLoadedForEditor() {
    if (!state.workbookArrayBuffer) {
      throw new ValidationError("Please load a spreadsheet file first (.xlsx, .ods, or .csv - upload or URL).");
    }
    const wb = readWorkbookFromArrayBuffer(state.workbookArrayBuffer);
    return wb;
  }

  function syncEditorUiFromState() {
    if (!els.editorScope || !els.editorSheet || !els.editorColumn) return;

    const ed = state.editor;
    els.editorScope.value = ed.scopeMode;

    // sheets
    els.editorSheet.disabled = ed.scopeMode !== "single" || !ed.workbookLoaded;
    if (ed.workbookLoaded) {
      const current = String(ed.selectedSheetName || "");
      els.editorSheet.innerHTML = "";
      const opt0 = document.createElement("option");
      opt0.value = "";
      opt0.textContent = "(Select sheet)";
      els.editorSheet.appendChild(opt0);
      for (const name of ed.workbookSheetNames || []) {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        els.editorSheet.appendChild(opt);
      }
      els.editorSheet.value = current && (ed.workbookSheetNames || []).includes(current) ? current : "";
    } else {
      els.editorSheet.innerHTML = `<option value="">(Load file first)</option>`;
    }

    // task
    if (els.editorTask) els.editorTask.value = ed.taskType;

    if (els.btnEditorDownload) els.btnEditorDownload.disabled = !Array.isArray(ed.previewRows);

    // preview sheet filter
    if (els.editorPreviewSheetFilter) {
      const sheetNames = ed.workbookSheetNames || [];
      els.editorPreviewSheetFilter.disabled = !Array.isArray(ed.previewRows);
      if (els.editorPreviewSheetFilter.options.length <= 1 && sheetNames.length) {
        els.editorPreviewSheetFilter.innerHTML = `<option value="">All Sheets</option>`;
        for (const name of sheetNames) {
          const opt = document.createElement("option");
          opt.value = name;
          opt.textContent = name;
          els.editorPreviewSheetFilter.appendChild(opt);
        }
      }
    }
  }

  function updateWizardUI() {
    const step = state.editor.wizardStep || 1;
    const ed = state.editor;

    // Update step indicators
    for (let i = 1; i <= 4; i++) {
      const stepEl = document.querySelector(`.wizard__step[data-step="${i}"]`);
      if (stepEl) {
        stepEl.classList.remove("is-active", "is-complete");
        if (i === step) {
          stepEl.classList.add("is-active");
        } else if (i < step) {
          stepEl.classList.add("is-complete");
        }
      }
    }

    // Show/hide panels
    for (let i = 1; i <= 4; i++) {
      const panelEl = document.querySelector(`.wizard__panel[data-panel="${i}"]`);
      if (panelEl) {
        panelEl.hidden = i !== step;
      }
    }

    // Update navigation buttons
    if (els.btnWizardPrev) {
      els.btnWizardPrev.disabled = step === 1;
    }
    if (els.btnWizardNext) {
      els.btnWizardNext.hidden = step === 4;
      els.btnWizardNext.disabled = !canProceedToNextStep(step);
    }
    if (els.btnWizardFinish) {
      els.btnWizardFinish.hidden = step !== 4;
      els.btnWizardFinish.disabled = !canProceedToNextStep(step);
    }

    // Update summary in step 4
    if (step === 4) {
      if (els.wizardSummaryFile) els.wizardSummaryFile.textContent = state.workbookName || "Not loaded";
      if (els.wizardSummaryMode) els.wizardSummaryMode.textContent = ed.scopeMode === "single" ? "Single-sheet" : "Multi-sheet";
      // For multi-sheet mode, always show "All sheets"; for single-sheet mode, show the selected sheet or "-"
      if (els.wizardSummarySheet) {
        els.wizardSummarySheet.textContent = ed.scopeMode === "multi" ? "All sheets" : (ed.selectedSheetName || "-");
      }
      // Get column header text - either from selectedColumn (if preview was generated) or look it up from selectedColumnKey
      if (els.wizardSummaryColumn) {
        let columnText = "-";
        if (ed.selectedColumn?.headerText) {
          columnText = ed.selectedColumn.headerText;
        } else if (ed.selectedColumnKey && ed.workbookLoaded) {
          try {
            const wb = ensureWorkbookLoadedForEditor();
            const scope = { mode: ed.scopeMode, sheetName: ed.selectedSheetName };
            const opts = listColumnOptions(wb, scope);
            const selected = opts.find((o) => o.key === ed.selectedColumnKey);
            if (selected) {
              columnText = selected.headerText;
            }
          } catch (e) {
            // If lookup fails, try parsing the key (format: "kind::headerText")
            const keyParts = String(ed.selectedColumnKey || "").split("::");
            if (keyParts.length >= 2) {
              columnText = keyParts.slice(1).join("::"); // In case headerText contains "::"
            }
          }
        }
        els.wizardSummaryColumn.textContent = columnText;
      }
      if (els.wizardSummaryTask) els.wizardSummaryTask.textContent = ed.taskType === "attendance" ? "Attendance input" : "Grade input";
      if (els.wizardSummaryInput) {
        // Show different text based on input method
        if (ed.inputMethod === "file") {
          els.wizardSummaryInput.textContent = ed.inputFileName || "-";
        } else {
          const lineCount = ed.inputTextContent ? ed.inputTextContent.split('\n').length : 0;
          els.wizardSummaryInput.textContent = lineCount > 0 ? `Text area (${lineCount} lines)` : "-";
        }
      }
    }
  }

  function canProceedToNextStep(step) {
    const ed = state.editor;
    switch (step) {
      case 1:
        // Enable Next if there's a file ready to load (uploaded or URL provided)
        return Boolean(state.workbookArrayBuffer) || Boolean(els.editorSheetUrl?.value?.trim());
      case 2:
        const needsSheet = ed.scopeMode === "single";
        return Boolean(ed.selectedColumnKey) && (!needsSheet || Boolean(ed.selectedSheetName));
      case 3:
        // Check if we have valid input based on input method
        if (ed.inputMethod === "file") {
          // File mode: check if a file is selected
          return Boolean(els.editorInputTxt?.files?.[0]);
        } else {
          // Textarea mode: check if textarea has content
          return Boolean(ed.inputTextContent && ed.inputTextContent.trim());
        }
      case 4:
        return ed.workbookLoaded && Boolean(ed.selectedColumnKey) && (
          (ed.inputMethod === "file" && Boolean(els.editorInputTxt?.files?.[0])) ||
          (ed.inputMethod === "textarea" && Boolean(ed.inputTextContent && ed.inputTextContent.trim()))
        );
      default:
        return false;
    }
  }

  function handleWizardPrev() {
    if (state.editor.wizardStep > 1) {
      state.editor.wizardStep--;
      updateWizardUI();
    }
  }

  async function handleWizardNext() {
    const step = state.editor.wizardStep || 1;
    
    // On step 1, automatically load the file before proceeding
    if (step === 1) {
      try {
        await handleEditorLoadFile();
        // Only proceed if loading was successful (workbookLoaded will be true)
        if (state.editor.workbookLoaded && canProceedToNextStep(step) && step < 4) {
          state.editor.wizardStep++;
          updateWizardUI();
        }
      } catch (e) {
        // Error already shown by handleEditorLoadFile, stay on step 1
        updateWizardUI();
      }
      return;
    }
    
    // For other steps, proceed normally
    if (canProceedToNextStep(step) && step < 4) {
      state.editor.wizardStep++;
      updateWizardUI();
    }
  }

  async function handleWizardFinish() {
    // This is step 4 - generate preview and switch to Reports
    await handleEditorBuildPreview();
  }

  function renderEditorPreview() {
    if (!els.editorPreviewTableBody) return;
    const ed = state.editor;
    const rows = Array.isArray(ed.previewRows) ? ed.previewRows : [];
    const mode = els.editorPreviewModeOrdered?.classList?.contains("is-active") ? "ordered" : "grouped";
    const sheetFilter = String(els.editorPreviewSheetFilter?.value || "");

    // Update summary
    if (els.summary) {
      if (!rows.length) {
        els.summary.textContent = "No preview generated. Go to Inputs tab to configure and generate preview.";
      } else {
        const total = rows.length;
        const matched = rows.filter((r) => r.match_status === "matched" || r.match_status === "manuallyFixed").length;
        const notFound = rows.filter((r) => r.match_status === "notFound").length;
        const ambiguous = rows.filter((r) => r.match_status === "ambiguous").length;
        els.summary.textContent = `Total: ${total} | Matched: ${matched} | Not Found: ${notFound} | Ambiguous: ${ambiguous}`;
      }
    }

    // simple ordering
    let out = rows.slice();
    if (sheetFilter) out = out.filter((r) => String(r.sheet || "") === sheetFilter);
    if (mode === "grouped") {
      out.sort((a, b) => {
        const s = String(a.sheet || "").localeCompare(String(b.sheet || ""));
        if (s !== 0) return s;
        return (a.row_index1 ?? 0) - (b.row_index1 ?? 0);
      });
    } else {
      out.sort((a, b) => (a.index ?? 0) - (b.index ?? 0));
    }

    els.editorPreviewTableBody.innerHTML = "";
    
    // In ordered mode, interleave delimiter rows if available
    if (mode === "ordered" && Array.isArray(ed.orderedEntries) && ed.orderedEntries.length > 0) {
      // Create a map of index -> preview row for quick lookup
      // The index in previewRows corresponds to the sequential position of IDs in orderedEntries
      const rowMap = new Map();
      for (const r of out) {
        if (r.index !== undefined && r.index !== null) {
          rowMap.set(Number(r.index), r);
        }
      }
      
      // Track which indices we've rendered
      const renderedIndices = new Set();
      let idCounter = 0; // Counter for IDs only (ignoring titles)
      let currentDelimiter = null; // Track current delimiter for data attributes
      
      for (const entry of ed.orderedEntries) {
        if (entry && typeof entry === "object") {
          if (entry.type === "title") {
            currentDelimiter = String(entry.title || "").trim();
            // Render delimiter row
            const tr = document.createElement("tr");
            tr.classList.add("row--delimiter");
            tr.dataset.delimiter = currentDelimiter;
            const td = document.createElement("td");
            td.colSpan = 9; // Merge all 9 columns (#, ID, Name, Sheet, Cell, Old, New, Status, Actions)
            td.textContent = currentDelimiter;
            td.className = "delimiter-cell";
            tr.appendChild(td);
            els.editorPreviewTableBody.appendChild(tr);
          } else if (entry.type === "id") {
            idCounter++;
            const row = rowMap.get(idCounter);
            if (row) {
              renderedIndices.add(idCounter);
              renderPreviewRow(row, currentDelimiter); // Pass delimiter to mark the row
            }
          }
        }
      }
      
      // Render any remaining rows that weren't in orderedEntries (shouldn't happen, but safety)
      for (const r of out) {
        if (r.index !== undefined && r.index !== null && !renderedIndices.has(Number(r.index))) {
          renderPreviewRow(r, null);
        }
      }
    } else {
      // Grouped mode or no orderedEntries - render normally
      for (const r of out) {
        renderPreviewRow(r, null);
      }
    }
    
    function renderPreviewRow(r, delimiter) {
      const tr = document.createElement("tr");
      if (r.match_status === "notFound" || r.match_status === "ambiguous") tr.classList.add("row--missing");
      if (r.discarded) tr.classList.add("row--discarded");
      // Check for duplicate IDs - use input_id (from input file) not student_id (from workbook match)
      if (ed.idCounts && r.input_id) {
        const normalizedId = normalizeId(r.input_id);
        const count = normalizedId ? (ed.idCounts[normalizedId] || 0) : 0;
        if (count > 1) {
          tr.classList.add("row--duplicate");
        }
      }
      if (delimiter) tr.dataset.delimiter = delimiter;
      const td = (t) => {
        const cell = document.createElement("td");
        cell.textContent = String(t ?? "");
        return cell;
      };
      tr.appendChild(td(r.index));
      tr.appendChild(td(r.student_id));
      tr.appendChild(td(r.student_name));
      tr.appendChild(td(r.sheet));
      tr.appendChild(td(r.cell));
      tr.appendChild(td(r.old_value));
      tr.appendChild(td(r.new_value));
      const statusTd = document.createElement("td");
      statusTd.textContent = String(r.match_status || "");
      tr.appendChild(statusTd);

      // actions
      const actionsTd = document.createElement("td");
      actionsTd.style.whiteSpace = "nowrap";
      actionsTd.style.display = "flex";
      actionsTd.style.gap = "8px";
      actionsTd.style.justifyContent = "center";

      const btnFix = document.createElement("button");
      btnFix.type = "button";
      btnFix.className = "btn btn--ghost";
      btnFix.textContent = "Fix";
      btnFix.dataset.action = "fix";
      btnFix.dataset.index = String(r.index);
      actionsTd.appendChild(btnFix);

      const btnMarkWrong = document.createElement("button");
      btnMarkWrong.type = "button";
      btnMarkWrong.className = "btn btn--ghost";
      btnMarkWrong.textContent = "Wrong";
      btnMarkWrong.dataset.action = "wrong";
      btnMarkWrong.dataset.index = String(r.index);
      actionsTd.appendChild(btnMarkWrong);

      const btnDiscard = document.createElement("button");
      btnDiscard.type = "button";
      btnDiscard.className = "btn btn--ghost";
      btnDiscard.textContent = r.discarded ? "Restore" : "Discard";
      btnDiscard.dataset.action = "discard";
      btnDiscard.dataset.index = String(r.index);
      actionsTd.appendChild(btnDiscard);

      if (String(state.editor.taskType || "") === "grade") {
        const btnEdit = document.createElement("button");
        btnEdit.type = "button";
        btnEdit.className = "btn btn--ghost";
        btnEdit.textContent = "Edit grade";
        btnEdit.dataset.action = "grade";
        btnEdit.dataset.index = String(r.index);
        actionsTd.appendChild(btnEdit);
      }

      tr.appendChild(actionsTd);
      els.editorPreviewTableBody.appendChild(tr);
    }
    
    // Populate delimiter filter dropdown and show/hide it based on mode
    if (els.delimiterFilterContainer && els.editorDelimiterFilter) {
      const isOrderedMode = mode === "ordered";
      const hasOrderedEntries = Array.isArray(ed.orderedEntries) && ed.orderedEntries.length > 0;
      
      if (isOrderedMode && hasOrderedEntries) {
        // Populate delimiter filter
        els.editorDelimiterFilter.innerHTML = '<option value="">Show All</option>';
        const delimiterSet = new Set();
        for (const entry of ed.orderedEntries) {
          if (entry && typeof entry === "object" && entry.type === "title") {
            const title = String(entry.title || "").trim();
            if (title && !delimiterSet.has(title)) {
              delimiterSet.add(title);
              const opt = document.createElement("option");
              opt.value = title;
              opt.textContent = title;
              els.editorDelimiterFilter.appendChild(opt);
            }
          }
        }
        els.delimiterFilterContainer.style.display = delimiterSet.size > 0 ? "block" : "none";
      } else {
        els.delimiterFilterContainer.style.display = "none";
      }
    }
    
    // Apply delimiter filter if one is selected
    applyDelimiterFilter();

    // final report mapping box
    if (els.editorFinalReportBox) {
      const map = Array.isArray(ed.columnMap) ? ed.columnMap : [];
      if (!map.length) {
        els.editorFinalReportBox.value = "(No mapping available)";
      } else {
        const lines = [];
        lines.push("MODIFIED COLUMN MAPPING");
        lines.push("=======================");
        lines.push(`Header: ${ed.selectedColumn?.headerText || ""}`);
        lines.push(`Kind: ${ed.selectedColumn?.kind || ""}`);
        lines.push("");
        for (const m of map) {
          lines.push(`- Sheet: ${m.sheet} | Header Row: ${m.header_row} | Column: ${m.col_letter}`);
        }
        els.editorFinalReportBox.value = lines.join("\n");
      }
    }
  }
  async function handleEditorXlsxUploadChanged() {
    const file = els.editorXlsxFile?.files?.[0] || null;
    if (!file) {
      state.workbookArrayBuffer = null;
      state.workbookName = null;
      state.editor.workbookLoaded = false;
      state.editor.workbookSheetNames = [];
      updateWizardUI();
      return;
    }
    try {
      state.workbookArrayBuffer = await readFileAsArrayBuffer(file);
      state.workbookName = file.name;
      // Reset workbookLoaded - it will be set to true when Next is clicked and file is processed
      state.editor.workbookLoaded = false;
      state.editor.workbookSheetNames = [];
      setEditorStatus(`File ready: ${file.name}`, "ok");
      updateWizardUI();
    } catch {
      state.workbookArrayBuffer = null;
      state.workbookName = null;
      state.editor.workbookLoaded = false;
      state.editor.workbookSheetNames = [];
      setEditorStatus("Failed to read the spreadsheet file.", "error");
      updateWizardUI();
    }
  }

  async function handleEditorLoadFile() {
    try {
      setEditorStatus("");
      disableRun(true, "Processing...");
      setEditorStatus("Loading workbook into memory…");

      // Use existing workbook buffer if upload already happened; otherwise attempt URL download.
      if (!state.workbookArrayBuffer) {
        const url = String(els.editorSheetUrl?.value || "").trim();
        if (!url) throw new ValidationError("Upload an .xlsx, .ods, or .csv file or provide a Google Sheet or SharePoint URL first.");
        try {
          const ab = await fetchXlsxFromUrl(url);
          state.workbookArrayBuffer = ab;
          state.workbookName = "downloaded_sheet.xlsx";
        } catch (downloadError) {
          // If download fails, throw a clear error message
          throw new DownloadError("The file could not be downloaded. Please check the URL or upload the file manually.");
        }
      }

      const wb = ensureWorkbookLoadedForEditor();
      state.editor.workbookLoaded = true;
      state.editor.workbookSheetNames = wb.SheetNames.slice();

      // initialize sheet selection if single mode
      if (state.editor.scopeMode === "single" && !state.editor.selectedSheetName) {
        state.editor.selectedSheetName = wb.SheetNames[0] || "";
      }

      // rebuild column list
      const scope = {
        mode: state.editor.scopeMode,
        sheetName: state.editor.selectedSheetName,
      };
      const opts = listColumnOptions(wb, scope);
      
      // Store column options for search functionality
      state.editor.columnOptions = opts;
      
      if (els.editorColumn) {
        els.editorColumn.disabled = false;
        els.editorColumn.innerHTML = "";
        const o0 = document.createElement("option");
        o0.value = "";
        o0.textContent = "(Select column)";
        els.editorColumn.appendChild(o0);
        for (const opt of opts) {
          const o = document.createElement("option");
          o.value = opt.key;
          const tag = opt.kind === "unknown" ? "Unknown" : opt.kind === "lecture" ? "Lecture" : "Section";
          o.textContent = `${opt.headerText} — ${tag} (${opt.occurrences})`;
          els.editorColumn.appendChild(o);
        }
      }
      
      // Enable column search
      if (els.editorColumnSearch) {
        els.editorColumnSearch.disabled = false;
        els.editorColumnSearch.value = "";
      }
      if (els.editorColumnSearchResults) {
        els.editorColumnSearchResults.innerHTML = "";
        els.editorColumnSearchResults.style.display = "none";
      }

      // enable sheet select if applicable
      if (els.editorSheet) els.editorSheet.disabled = state.editor.scopeMode !== "single";
      setEditorStatus("Workbook loaded. Select mode/sheet/column and upload the input file.", "ok");
      syncEditorUiFromState();
      
      // Don't auto-advance here - let handleWizardNext handle it
      updateWizardUI();
    } catch (e) {
      const msg =
        e instanceof ValidationError || e instanceof DownloadError || e instanceof FileError || e instanceof ProcessingError
          ? e.message
          : `Unexpected error: ${e?.message || String(e)}`;
      setEditorStatus(msg, "error");
      // Reset workbook state on error
      state.workbookArrayBuffer = null;
      state.workbookName = null;
      state.editor.workbookLoaded = false;
      state.editor.workbookSheetNames = [];
      throw e; // Re-throw so handleWizardNext knows it failed
    } finally {
      disableRun(false, "Run Attendance Check");
    }
  }

  function handleEditorSelectionChanged() {
    try {
      const oldScopeMode = state.editor.scopeMode;
      if (els.editorScope) state.editor.scopeMode = String(els.editorScope.value || "single");
      
      // Clear selectedSheetName when switching to multi-sheet mode
      if (state.editor.scopeMode === "multi" && oldScopeMode === "single") {
        state.editor.selectedSheetName = "";
      }
      
      if (els.editorSheet) state.editor.selectedSheetName = String(els.editorSheet.value || "");
      if (els.editorColumn) {
        state.editor.selectedColumnKey = String(els.editorColumn.value || "");
        // Clear search when column is selected from dropdown
        if (els.editorColumnSearch) {
          els.editorColumnSearch.value = "";
        }
        if (els.editorColumnSearchResults) {
          els.editorColumnSearchResults.innerHTML = "";
          els.editorColumnSearchResults.style.display = "none";
        }
      }

      // Rebuild columns if workbook is loaded
      if (state.editor.workbookLoaded) {
        const wb = ensureWorkbookLoadedForEditor();
        const scope = { mode: state.editor.scopeMode, sheetName: state.editor.selectedSheetName };
        const opts = listColumnOptions(wb, scope);
        
        // Store column options for search functionality
        state.editor.columnOptions = opts;
        
        if (els.editorColumn) {
          const cur = state.editor.selectedColumnKey;
          els.editorColumn.innerHTML = "";
          const o0 = document.createElement("option");
          o0.value = "";
          o0.textContent = "(Select column)";
          els.editorColumn.appendChild(o0);
          for (const opt of opts) {
            const o = document.createElement("option");
            o.value = opt.key;
            const tag = opt.kind === "unknown" ? "Unknown" : opt.kind === "lecture" ? "Lecture" : "Section";
            o.textContent = `${opt.headerText} — ${tag} (${opt.occurrences})`;
            els.editorColumn.appendChild(o);
          }
          els.editorColumn.value = opts.some((x) => x.key === cur) ? cur : "";
          state.editor.selectedColumnKey = String(els.editorColumn.value || "");
        }
        
        // Enable column search and clear previous search when mode/sheet changes
        if (els.editorColumnSearch) {
          els.editorColumnSearch.disabled = false;
          els.editorColumnSearch.value = "";
        }
        if (els.editorColumnSearchResults) {
          els.editorColumnSearchResults.innerHTML = "";
          els.editorColumnSearchResults.style.display = "none";
        }
      } else {
        // Disable column search if workbook not loaded
        if (els.editorColumnSearch) {
          els.editorColumnSearch.disabled = true;
          els.editorColumnSearch.value = "";
        }
        if (els.editorColumnSearchResults) {
          els.editorColumnSearchResults.innerHTML = "";
          els.editorColumnSearchResults.style.display = "none";
        }
      }
      syncEditorUiFromState();
      updateWizardUI();
    } catch (e) {
      setEditorStatus(e?.message || String(e), "error");
    }
  }

  function handleEditorTaskChanged() {
    state.editor.taskType = String(els.editorTask?.value || "attendance");
    syncEditorUiFromState();
  }

  function handleEditorInputMethodChanged() {
    const isFile = els.editorInputMethodFile?.checked;
    state.editor.inputMethod = isFile ? "file" : "textarea";
    
    // Show/hide appropriate input containers
    if (els.editorInputFileContainer) {
      els.editorInputFileContainer.style.display = isFile ? "block" : "none";
    }
    if (els.editorInputTextareaContainer) {
      els.editorInputTextareaContainer.style.display = isFile ? "none" : "block";
    }
    
    // Update wizard UI to check if we can proceed
    updateWizardUI();
  }

  function handleEditorInputChanged() {
    const f = els.editorInputTxt?.files?.[0] || null;
    state.editor.inputFileName = f ? f.name : "";
    syncEditorUiFromState();
    updateWizardUI();
  }

  function handleEditorTextareaChanged() {
    const textarea = els.editorInputTextarea;
    if (!textarea) return;
    
    const content = String(textarea.value || "").trim();
    state.editor.inputTextContent = content || null;
    
    updateWizardUI();
  }

  async function handleEditorBuildPreview() {
    try {
      setEditorStatus("");
      disableRun(true, "Processing...");
      setEditorStatus("Building preview…");

      const wb = ensureWorkbookLoadedForEditor();
      const ed = state.editor;

      const scope = { mode: ed.scopeMode, sheetName: ed.selectedSheetName };
      const colKey = String(ed.selectedColumnKey || "").trim();
      if (!colKey) throw new ValidationError("Please select a target column.");
      if (ed.scopeMode === "single" && !ed.selectedSheetName) throw new ValidationError("Please select a sheet.");

      // Get input text - either from file or textarea
      let inputText;
      if (ed.inputMethod === "file") {
        const inputFile = els.editorInputTxt?.files?.[0] || null;
        if (!inputFile) throw new ValidationError("Input .txt file is required.");
        inputText = await readFileAsText(inputFile);
      } else {
        // textarea mode
        inputText = ed.inputTextContent;
        if (!inputText || !inputText.trim()) {
          throw new ValidationError("Please enter input data in the text area.");
        }
      }

      const task = String(ed.taskType || "attendance");

      let preview;
      let orderedEntries = null;
      let idCounts = null;
      if (task === "attendance") {
        const parsed = parseStudentIdsText(inputText);
        orderedEntries = parsed.orderedEntries; // Store for delimiter rendering
        idCounts = parsed.idCounts; // Extract idCounts for duplicate detection
        const orderedIds = parsed.orderedEntries
          .filter((x) => x && typeof x === "object" && x.type === "id")
          .map((x) => String(x.id));
        preview = computeEditorPreview({
          workbook: wb,
          scope,
          columnKey: colKey,
          taskType: "attendance",
          orderedAttendanceIds: orderedIds,
          attendanceIdsSet: parsed.targetIdsSet,
          gradesRows: null,
        });
        
        // Store original parsed data for download functionality
        ed.originalInputData = {
          type: "attendance",
          orderedEntries: parsed.orderedEntries,
          idsSet: parsed.targetIdsSet,
        };
      } else {
        const parsed = parseGradesText(inputText);
        orderedEntries = parsed.orderedEntries; // Store for delimiter rendering
        // Calculate idCounts from orderedEntries for grades - normalize IDs for consistent matching
        idCounts = {};
        for (const entry of parsed.orderedEntries || []) {
          if (entry && typeof entry === "object" && entry.type === "id" && entry.id) {
            const sid = normalizeId(entry.id);
            if (sid) {
              idCounts[sid] = (idCounts[sid] || 0) + 1;
            }
          }
        }
        preview = computeEditorPreview({
          workbook: wb,
          scope,
          columnKey: colKey,
          taskType: "grade",
          orderedAttendanceIds: null,
          attendanceIdsSet: null,
          gradesRows: parsed.rows, // Use rows array from parsed result
        });
        
        // Store original parsed data for download functionality
        ed.originalInputData = {
          type: "grade",
          orderedEntries: parsed.orderedEntries,
          rows: parsed.rows,
        };
      }

      ed.previewRows = preview.preview_rows;
      ed.columnMap = preview.column_map;
      ed.selectedColumn = preview.selected_column;
      ed.orderedEntries = orderedEntries; // Store delimiter information (works for both attendance and grades)
      ed.idCounts = idCounts; // Store idCounts for duplicate highlighting
      editorSearchRows = buildStudentSearchRows(wb, scope);

      // enable preview sheet filter
      if (els.editorPreviewSheetFilter) {
        els.editorPreviewSheetFilter.disabled = false;
        // populate sheet filter options
        const sheetSet = new Set();
        for (const r of ed.previewRows || []) {
          const s = String(r.sheet || "").trim();
          if (s) sheetSet.add(s);
        }
        const sheets = Array.from(sheetSet).sort();
        els.editorPreviewSheetFilter.innerHTML = '<option value="">All Sheets</option>';
        for (const s of sheets) {
          const opt = document.createElement("option");
          opt.value = s;
          opt.textContent = s;
          els.editorPreviewSheetFilter.appendChild(opt);
        }
      }

      // enable download buttons
      if (els.btnEditorDownload) els.btnEditorDownload.disabled = false;
      if (els.btnDownloadModifiedRecords) els.btnDownloadModifiedRecords.disabled = false;
      if (els.btnDownloadOriginalRecords) els.btnDownloadOriginalRecords.disabled = false;
      if (els.btnDownloadJson) els.btnDownloadJson.disabled = false;
      if (els.btnDownloadTxt) els.btnDownloadTxt.disabled = false;
      if (els.btnDownloadPdf) els.btnDownloadPdf.disabled = false;

      renderEditorPreview();
      syncEditorUiFromState();
      setEditorStatus("Preview generated. Review carefully, then download when ready.", "ok");
      
      // Auto-switch to Reports view
      switchView("report");
    } catch (e) {
      const msg =
        e instanceof ValidationError || e instanceof FileError || e instanceof ProcessingError || e instanceof DownloadError
          ? e.message
          : `Unexpected error: ${e?.message || String(e)}`;
      setEditorStatus(msg, "error");
    } finally {
      disableRun(false, "Run Attendance Check");
    }
  }

  function handleEditorPreviewModeChanged() {
    // toggle button active state
    const grouped = els.editorPreviewModeGrouped;
    const ordered = els.editorPreviewModeOrdered;
    if (grouped && ordered) {
      const clickedGrouped = document.activeElement === grouped;
      const clickedOrdered = document.activeElement === ordered;
      if (clickedGrouped) {
        grouped.classList.add("is-active");
        grouped.setAttribute("aria-selected", "true");
        ordered.classList.remove("is-active");
        ordered.setAttribute("aria-selected", "false");
      } else if (clickedOrdered) {
        ordered.classList.add("is-active");
        ordered.setAttribute("aria-selected", "true");
        grouped.classList.remove("is-active");
        grouped.setAttribute("aria-selected", "false");
      }
    }
    renderEditorPreview();
  }

  function openFixDialogForIndex(idx) {
    editorActiveFixIndex = idx;
    if (!els.editorFixDialog || !els.editorFixSearch || !els.editorFixResults) return;
    
    // Ensure search rows are populated - rebuild if empty or if workbook is available
    if (!editorSearchRows || editorSearchRows.length === 0) {
      try {
        const wb = ensureWorkbookLoadedForEditor();
        const ed = state.editor;
        const scope = { mode: ed.scopeMode, sheetName: ed.selectedSheetName };
        editorSearchRows = buildStudentSearchRows(wb, scope);
      } catch (e) {
        console.error("Failed to build search rows:", e);
        editorSearchRows = [];
      }
    }
    
    els.editorFixSearch.value = "";
    els.editorFixResults.innerHTML = "";
    els.editorFixDialog.showModal();
    els.editorFixSearch.focus();
  }

  function openGradeDialogForIndex(idx) {
    editorActiveGradeIndex = idx;
    if (!els.editorGradeDialog || !els.editorGradeValue) return;
    const row = (state.editor.previewRows || []).find((r) => Number(r.index) === Number(idx));
    els.editorGradeValue.value = row ? String(row.new_value ?? "") : "";
    els.editorGradeDialog.showModal();
    els.editorGradeValue.focus();
  }

  function handleEditorPreviewRowAction(e) {
    const target = e?.target;
    const btn = target?.closest?.("button[data-action]");
    if (!btn) return;
    const action = String(btn.dataset.action || "");
    const idx = Number.parseInt(String(btn.dataset.index || ""), 10);
    if (!Number.isFinite(idx)) return;
    if (!Array.isArray(state.editor.previewRows)) return;

    if (action === "fix") {
      openFixDialogForIndex(idx);
      return;
    }
    if (action === "grade") {
      openGradeDialogForIndex(idx);
      return;
    }
    if (action === "wrong") {
      const row = state.editor.previewRows.find((r) => Number(r.index) === idx);
      if (row) {
        row.match_status = "ambiguous";
        row.note = "Marked as wrong match by user.";
        renderEditorPreview();
      }
    }
    if (action === "discard") {
      const row = state.editor.previewRows.find((r) => Number(r.index) === idx);
      if (row) {
        // Toggle discarded state
        row.discarded = !row.discarded;
        renderEditorPreview();
      }
    }
  }

  function handleEditorFixSearchChanged() {
    if (!els.editorFixSearch || !els.editorFixResults) return;
    const q = String(els.editorFixSearch.value || "").trim().toLowerCase();
    els.editorFixResults.innerHTML = "";
    if (!q) return;

    // Ensure search rows are available
    if (!editorSearchRows || editorSearchRows.length === 0) {
      try {
        const wb = ensureWorkbookLoadedForEditor();
        const ed = state.editor;
        const scope = { mode: ed.scopeMode, sheetName: ed.selectedSheetName };
        editorSearchRows = buildStudentSearchRows(wb, scope);
      } catch (e) {
        console.error("Failed to build search rows:", e);
        els.editorFixResults.innerHTML = '<div class="searchResults__empty">Unable to load student data. Please regenerate the preview.</div>';
        return;
      }
    }

    const results = [];
    for (const r of editorSearchRows || []) {
      const id = String(r.id || "");
      const name = String(r.name || "");
      if (id.toLowerCase().includes(q) || name.toLowerCase().includes(q)) results.push(r);
      if (results.length >= 30) break;
    }

    if (results.length === 0) {
      els.editorFixResults.innerHTML = '<div class="searchResults__empty">No students found matching your search.</div>';
      return;
    }

    for (const r of results) {
      const item = document.createElement("div");
      item.className = "searchResults__item";
      item.dataset.sheet = r.sheet;
      item.dataset.row1 = String(r.row1);
      item.dataset.id = r.id;
      item.dataset.name = r.name;

      const main = document.createElement("div");
      main.className = "searchResults__main";
      const line1 = document.createElement("div");
      line1.innerHTML = `<span class="searchResults__id">${escapeHtml(r.id)}</span> — ${escapeHtml(r.name)}`;
      const meta = document.createElement("div");
      meta.className = "searchResults__meta";
      meta.textContent = `Sheet: ${r.sheet} | Row: ${r.row1}`;
      main.appendChild(line1);
      main.appendChild(meta);

      item.appendChild(main);
      els.editorFixResults.appendChild(item);
    }
  }

  function handleEditorFixResultClicked(e) {
    const node = e?.target?.closest?.(".searchResults__item");
    if (!node) return;
    if (!Number.isFinite(editorActiveFixIndex)) return;
    const idx = Number(editorActiveFixIndex);
    const sheet = String(node.dataset.sheet || "");
    const row1 = Number.parseInt(String(node.dataset.row1 || ""), 10);
    const id = String(node.dataset.id || "");
    const name = String(node.dataset.name || "");
    if (!sheet || !Number.isFinite(row1)) return;

    const wb = ensureWorkbookLoadedForEditor();
    const ws = wb.Sheets[sheet];
    if (!ws) return;

    // find column location for this sheet
    const colLoc = (state.editor.selectedColumn?.locations || []).find((l) => String(l.sheet) === sheet);
    if (!colLoc) return;

    const addr = window.XLSX.utils.encode_cell({ r: row1 - 1, c: colLoc.col1 - 1 });
    const cell = ws[addr];
    const oldVal = cell?.v ?? "";

    const row = state.editor.previewRows.find((r) => Number(r.index) === idx);
    if (row) {
      row.sheet = sheet;
      row.row_index1 = row1;
      row.student_id = id;
      row.student_name = name;
      row.col_letter = window.XLSX.utils.encode_col(colLoc.col1 - 1);
      row.cell = addr;
      row.old_value = oldVal === null || oldVal === undefined ? "" : oldVal;
      row.match_status = "manuallyFixed";
      row.note = "Manually fixed by user.";
    }

    if (els.editorFixDialog) els.editorFixDialog.close();
    editorActiveFixIndex = null;
    renderEditorPreview();
  }

  function handleEditorGradeSaveClicked() {
    if (!Number.isFinite(editorActiveGradeIndex)) return;
    const idx = Number(editorActiveGradeIndex);
    const v = String(els.editorGradeValue?.value ?? "").trim();
    const row = state.editor.previewRows?.find((r) => Number(r.index) === idx);
    if (row) row.new_value = v;
    editorActiveGradeIndex = null;
    if (els.editorGradeDialog) els.editorGradeDialog.close();
    renderEditorPreview();
  }

  function escapeHtml(s) {
    return String(s ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function handleEditorDownloadModified() {
    try {
      const ed = state.editor;
      const rows = Array.isArray(ed.previewRows) ? ed.previewRows : [];
      if (!rows.length) throw new ValidationError("Please generate a preview first.");

      const warning =
        "High-responsibility operation.\n\n" +
        "You are about to generate and download a modified Excel file.\n" +
        "Strongly recommended: review the preview carefully (IDs, names, sheet, cell, old/new).\n\n" +
        "Do you want to continue?";
      const ok = window.confirm(warning);
      if (!ok) return;

      setEditorStatus("Generating modified workbook…");

      const wb = ensureWorkbookLoadedForEditor();
      
      // Read highlight settings from state
      const highlightEnabled = ed.highlightEnabled ?? true;
      const highlightColor = ed.highlightColor ?? "#FFFF00";
      
      // Filter out discarded records before applying edits
      const activeRows = rows.filter((r) => !r.discarded);
      
      // Apply edits with highlight settings
      applyEditorEdits(wb, activeRows, highlightEnabled, highlightColor);

      if (!window.XLSX || !window.XLSX.write) {
        throw new ProcessingError("XLSX writer not loaded. Please refresh the page.");
      }
      // Include cellStyles option to write cell styling information
      const out = window.XLSX.write(wb, { 
        bookType: "xlsx", 
        type: "array",
        cellStyles: true
      });

      const base = safeBaseName(state.workbookName || "workbook");
      const filename = `${base}_modified.xlsx`;
      downloadBlob(
        filename,
        out,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      // Final report is already rendered in the textarea; keep it visible and update status.
      const highlightMsg = highlightEnabled ? " (with highlighted cells)" : "";
      const discardedCount = rows.length - activeRows.length;
      const discardedMsg = discardedCount > 0 ? ` ${discardedCount} discarded record(s) were excluded.` : "";
      setEditorStatus(`Downloaded modified file${highlightMsg}.${discardedMsg} Final column mapping report is shown below.`, "ok");
    } catch (e) {
      const msg =
        e instanceof ValidationError || e instanceof FileError || e instanceof ProcessingError
          ? e.message
          : `Unexpected error: ${e?.message || String(e)}`;
      setEditorStatus(msg, "error");
    }
  }

  function handleDownloadJson() {
    // Check for editor preview first, then fall back to legacy report
    const ed = state.editor;
    const rows = Array.isArray(ed.previewRows) ? ed.previewRows : [];
    if (rows.length) {
      const exportData = {
        metadata: {
          timestamp: new Date().toISOString(),
          task_type: ed.taskType || "attendance",
          column_header: ed.selectedColumn?.headerText || "",
          column_kind: ed.selectedColumn?.kind || "",
          scope_mode: ed.scopeMode || "",
          selected_sheet: ed.selectedSheetName || "",
        },
        preview_rows: rows,
        column_map: ed.columnMap || [],
        selected_column: ed.selectedColumn || null,
        ordered_entries: ed.orderedEntries || null, // Include delimiter information
      };
      const base = safeBaseName(state.workbookName || "workbook");
      const json = JSON.stringify(exportData, null, 2);
      downloadBlob(`${base}_preview_report.json`, json, "application/json;charset=utf-8");
      return;
    }
    // No preview data available
    if (els.summary) {
      els.summary.textContent = "No preview generated. Generate a preview first.";
    }
  }

  function handleDownloadTxt() {
    // Check for editor preview first, then fall back to legacy report
    const ed = state.editor;
    const rows = Array.isArray(ed.previewRows) ? ed.previewRows : [];
    if (rows.length) {
      const lines = [];
      lines.push("EDIT PREVIEW REPORT");
      lines.push("===================");
      lines.push(`Task Type: ${ed.taskType || "attendance"}`);
      lines.push(`Column Header: ${ed.selectedColumn?.headerText || ""}`);
      lines.push(`Column Kind: ${ed.selectedColumn?.kind || ""}`);
      lines.push(`Scope: ${ed.scopeMode || ""}${ed.selectedSheetName ? ` (Sheet: ${ed.selectedSheetName})` : ""}`);
      lines.push(`Generated: ${new Date().toISOString()}`);
      lines.push("");
      lines.push("PREVIEW ROWS");
      lines.push("-".repeat(60));
      lines.push("");
      for (const r of rows) {
        lines.push(`Row ${r.index}: ${r.student_id} | ${r.student_name || "N/A"} | Sheet: ${r.sheet || "N/A"} | Cell: ${r.cell || "N/A"}`);
        lines.push(`  Old: ${r.old_value || ""} → New: ${r.new_value || ""} | Status: ${r.match_status || ""}`);
        if (r.note) lines.push(`  Note: ${r.note}`);
        lines.push("");
      }
      lines.push("COLUMN MAPPING");
      lines.push("-".repeat(60));
      for (const m of ed.columnMap || []) {
        lines.push(`Sheet: ${m.sheet} | Header Row: ${m.header_row} | Column: ${m.col_letter}`);
      }
      const base = safeBaseName(state.workbookName || "workbook");
      const txt = lines.join("\n");
      downloadBlob(`${base}_preview_report.txt`, txt, "text/plain;charset=utf-8");
      return;
    }
    // No preview data available
    if (els.summary) {
      els.summary.textContent = "No preview generated. Generate a preview first.";
    }
  }

  async function handleLoadPreviousReportJson() {
    const file = els.loadReportJson.files?.[0] || null;
    if (!file) return;
    try {
      const txt = await readFileAsText(file);
      const parsed = JSON.parse(txt);
      
      // Only support editor preview format
      if (parsed.preview_rows && Array.isArray(parsed.preview_rows)) {
        state.editor.previewRows = parsed.preview_rows;
        state.editor.columnMap = parsed.column_map || [];
        state.editor.selectedColumn = parsed.selected_column || null;
        state.editor.orderedEntries = parsed.ordered_entries || null; // Preserve delimiter information
        if (parsed.metadata) {
          state.editor.taskType = parsed.metadata.task_type || "attendance";
          state.editor.scopeMode = parsed.metadata.scope_mode || "single";
          state.editor.selectedSheetName = parsed.metadata.selected_sheet || "";
        }
        
        // enable preview sheet filter and populate options
        if (els.editorPreviewSheetFilter) {
          els.editorPreviewSheetFilter.disabled = false;
          // populate sheet filter options from preview rows
          const sheetSet = new Set();
          for (const r of state.editor.previewRows || []) {
            const s = String(r.sheet || "").trim();
            if (s) sheetSet.add(s);
          }
          const sheets = Array.from(sheetSet).sort();
          els.editorPreviewSheetFilter.innerHTML = '<option value="">All Sheets</option>';
          for (const s of sheets) {
            const opt = document.createElement("option");
            opt.value = s;
            opt.textContent = s;
            els.editorPreviewSheetFilter.appendChild(opt);
          }
        }
        
        renderEditorPreview();
        els.btnEditorDownload.disabled = false;
        els.btnDownloadJson.disabled = false;
        els.btnDownloadTxt.disabled = false;
        els.btnDownloadPdf.disabled = false;
        if (els.summary) {
          els.summary.textContent = `Loaded preview from: ${file.name}`;
        }
        switchView("report");
      } else {
        throw new Error("Invalid format. Expected editor preview JSON with 'preview_rows' array.");
      }
    } catch (e) {
      if (els.summary) {
        els.summary.textContent = `Failed to load report JSON: ${e?.message || String(e)}`;
      }
    } finally {
      els.loadReportJson.value = "";
    }
  }

  function assertPdfLibs() {
    const jspdf = window.jspdf;
    if (!jspdf || !jspdf.jsPDF) {
      throw new Error("PDF library not loaded. Please refresh the page.");
    }
    // autotable attaches to jsPDF prototype; we'll check for method existence
    const doc = new jspdf.jsPDF();
    if (typeof doc.autoTable !== "function") {
      throw new Error("PDF table plugin not loaded. Please refresh the page.");
    }
  }

  function handleDownloadPdf() {
    assertPdfLibs();
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });

    // Check for editor preview first
    const ed = state.editor;
    const previewRows = Array.isArray(ed.previewRows) ? ed.previewRows : [];
    if (previewRows.length) {
      const title = "Edit Preview Report";
      const subtitleParts = [
        `Task: ${ed.taskType || "attendance"}`,
        `Column: ${ed.selectedColumn?.headerText || ""}`,
        `Generated: ${new Date().toISOString()}`,
      ].filter(Boolean);

      const sheetFilter = String(els.editorPreviewSheetFilter?.value || "");
      const headerLeft = title;
      const headerRight = sheetFilter ? `Sheet: ${sheetFilter}` : "All Sheets";

      // Header
      doc.setFont("helvetica", "bold");
      doc.setFontSize(18);
      doc.text(headerLeft, 40, 40);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(11);
      doc.text(headerRight, doc.internal.pageSize.getWidth() - 40, 40, { align: "right" });

      doc.setFontSize(11);
      doc.setTextColor(90);
      doc.text(subtitleParts.join(" | "), 40, 62);
      doc.setTextColor(0);

      // Build table rows
      let out = previewRows.slice();
      if (sheetFilter) out = out.filter((r) => String(r.sheet || "") === sheetFilter);
      const mode = els.editorPreviewModeOrdered?.classList?.contains("is-active") ? "ordered" : "grouped";
      if (mode === "grouped") {
        out.sort((a, b) => {
          const s = String(a.sheet || "").localeCompare(String(b.sheet || ""));
          if (s !== 0) return s;
          return (a.row_index1 ?? 0) - (b.row_index1 ?? 0);
        });
      } else {
        out.sort((a, b) => (a.index ?? 0) - (b.index ?? 0));
      }

      const pdfRows = out.map((r) => [
        String(r.index || ""),
        String(r.student_id || ""),
        String(r.student_name || ""),
        String(r.sheet || ""),
        String(r.cell || ""),
        String(r.old_value || ""),
        String(r.new_value || ""),
        String(r.match_status || ""),
      ]);

      doc.autoTable({
        startY: 80,
        head: [["#", "ID", "Name", "Sheet", "Cell", "Old", "New", "Status"]],
        body: pdfRows,
        styles: {
          font: "helvetica",
          fontSize: 8,
          cellPadding: 4,
          overflow: "linebreak",
          valign: "top",
        },
        headStyles: {
          fillColor: [28, 36, 57],
          textColor: 255,
          fontStyle: "bold",
        },
        alternateRowStyles: {
          fillColor: [245, 247, 255],
        },
        columnStyles: {
          0: { cellWidth: 30 },
          1: { cellWidth: 70 },
          2: { cellWidth: 150 },
          3: { cellWidth: 80 },
          4: { cellWidth: 60 },
          5: { cellWidth: 50 },
          6: { cellWidth: 50 },
          7: { cellWidth: 70 },
        },
        didParseCell: (data) => {
          if (data.section !== "body") return;
          const row = out[data.row.index];
          if (row && (row.match_status === "notFound" || row.match_status === "ambiguous")) {
            data.cell.styles.fillColor = [255, 235, 235];
            data.cell.styles.textColor = [160, 25, 25];
          }
        },
        didDrawPage: (data) => {
          const pageCount = doc.internal.getNumberOfPages();
          const pageSize = doc.internal.pageSize;
          doc.setFontSize(9);
          doc.setTextColor(120);
          doc.text(`Page ${data.pageNumber} of ${pageCount}`, pageSize.getWidth() - 40, pageSize.getHeight() - 20, {
            align: "right",
          });
          doc.setTextColor(0);
        },
      });

      const base = safeBaseName(state.workbookName || "workbook");
      doc.save(`${base}_preview_report.pdf`);
      return;
    }

    // No legacy report support - only editor preview
    if (!previewRows.length) {
      if (els.summary) {
        els.summary.textContent = "No preview generated. Generate a preview first.";
      }
      return;
    }
  }

  function handleDownloadModifiedRecords() {
    try {
      const ed = state.editor;
      const rows = Array.isArray(ed.previewRows) ? ed.previewRows : [];
      if (!rows.length) {
        throw new ValidationError("Please generate a preview first.");
      }

      const taskType = ed.taskType || "attendance";
      const lines = [];

      if (taskType === "attendance") {
        // For attendance: output only matched IDs from preview (excluding discarded)
        // Preserve delimiter structure from orderedEntries
        if (Array.isArray(ed.orderedEntries) && ed.orderedEntries.length > 0) {
          // Build a map of ID -> preview row for quick lookup
          const rowMap = new Map();
          for (const r of rows) {
            if (r.student_id && !r.discarded) {
              rowMap.set(r.student_id, r);
            }
          }

          // Iterate through ordered entries and output modified data
          for (const entry of ed.orderedEntries) {
            if (entry && typeof entry === "object") {
              if (entry.type === "title") {
                lines.push(entry.title || "");
              } else if (entry.type === "id") {
                const previewRow = rowMap.get(entry.id);
                // Only include if found in preview AND not discarded
                if (previewRow) {
                  // Use the student_id from preview (in case it was manually fixed)
                  lines.push(previewRow.student_id);
                }
              }
            }
          }
        } else {
          // Fallback: just list IDs from preview rows (excluding discarded)
          for (const r of rows) {
            if (r.student_id && !r.discarded) {
              lines.push(r.student_id);
            }
          }
        }
      } else {
        // For grades: preserve delimiter structure and output id,grade from preview
        if (Array.isArray(ed.orderedEntries) && ed.orderedEntries.length > 0) {
          // Build a map of ID -> preview row for quick lookup
          const rowMap = new Map();
          for (const r of rows) {
            if (r.student_id && !r.discarded) {
              rowMap.set(r.input_id || r.student_id, r);
            }
          }

          // Iterate through ordered entries and output modified data
          for (const entry of ed.orderedEntries) {
            if (entry && typeof entry === "object") {
              if (entry.type === "title") {
                lines.push(entry.title || "");
              } else if (entry.type === "id") {
                const previewRow = rowMap.get(entry.id);
                // Only include if found in preview AND not discarded
                if (previewRow) {
                  // Use the modified student_id and new_value from preview
                  const grade = String(previewRow.new_value || "");
                  lines.push(`${previewRow.student_id},${grade}`);
                }
              }
            }
          }
        } else {
          // Fallback: just list id,grade from preview rows (excluding discarded)
          for (const r of rows) {
            if (r.student_id && !r.discarded) {
              lines.push(`${r.student_id},${r.new_value || ""}`);
            }
          }
        }
      }

      const txtContent = lines.join("\n");
      const base = safeBaseName(state.workbookName || "workbook");
      const filename = `${base}_modified_records.txt`;
      downloadBlob(filename, txtContent, "text/plain;charset=utf-8");

      setEditorStatus("Downloaded modified records as text file.", "ok");
    } catch (e) {
      const msg = e instanceof ValidationError ? e.message : `Error: ${e?.message || String(e)}`;
      setEditorStatus(msg, "error");
    }
  }

  function handleDownloadOriginalRecords() {
    try {
      const ed = state.editor;
      const originalData = ed.originalInputData;
      if (!originalData) {
        throw new ValidationError("No original input data available. Please generate a preview first.");
      }

      const lines = [];

      if (originalData.type === "attendance") {
        // For attendance: reconstruct from orderedEntries
        if (Array.isArray(originalData.orderedEntries) && originalData.orderedEntries.length > 0) {
          for (const entry of originalData.orderedEntries) {
            if (entry && typeof entry === "object") {
              if (entry.type === "title") {
                lines.push(entry.title || "");
              } else if (entry.type === "id") {
                lines.push(entry.id);
              }
            }
          }
        }
      } else if (originalData.type === "grade") {
        // For grades: reconstruct from orderedEntries
        if (Array.isArray(originalData.orderedEntries) && originalData.orderedEntries.length > 0) {
          for (const entry of originalData.orderedEntries) {
            if (entry && typeof entry === "object") {
              if (entry.type === "title") {
                lines.push(entry.title || "");
              } else if (entry.type === "id") {
                lines.push(`${entry.id},${entry.grade || ""}`);
              }
            }
          }
        }
      }

      if (lines.length === 0) {
        throw new ValidationError("No original data to download.");
      }

      const txtContent = lines.join("\n");
      const base = safeBaseName(state.workbookName || "workbook");
      const filename = `${base}_original_records.txt`;
      downloadBlob(filename, txtContent, "text/plain;charset=utf-8");

      setEditorStatus("Downloaded original records as text file.", "ok");
    } catch (e) {
      const msg = e instanceof ValidationError ? e.message : `Error: ${e?.message || String(e)}`;
      setEditorStatus(msg, "error");
    }
  }

  return {
    handleDownloadJson,
    handleDownloadTxt,
    handleLoadPreviousReportJson,
    handleDownloadPdf,
    handleDownloadModifiedRecords,
    handleDownloadOriginalRecords,

    // Editor handlers
    handleEditorXlsxUploadChanged,
    handleEditorLoadFile,
    handleEditorSelectionChanged,
    handleEditorTaskChanged,
    handleEditorInputMethodChanged,
    handleEditorInputChanged,
    handleEditorTextareaChanged,
    handleEditorBuildPreview,
    handleEditorPreviewModeChanged,
    handleEditorDownloadModified,

    // Preview fixups
    handleEditorPreviewRowAction,
    handleEditorFixSearchChanged,
    handleEditorFixResultClicked,
    handleEditorGradeSaveClicked,
    handleDelimiterFilterChanged,

    // Wizard navigation
    handleWizardPrev,
    handleWizardNext,
    handleWizardFinish,
    updateWizardUI,
    
    // Column search
    handleEditorColumnSearchChanged,
    handleEditorColumnSearchResultClicked,
    
    // OCR handlers
    handleOcrImageUpload,
    handleOcrProcess,
    handleOcrApprove,
    handleOcrReset,
    updateOcrUI,

    // Sheet Merger handlers
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
  
  function applyDelimiterFilter() {
    if (!els.editorPreviewTableBody || !els.editorDelimiterFilter) return;
    const selectedDelimiter = String(els.editorDelimiterFilter.value || "").trim();
    const mode = els.editorPreviewModeOrdered?.classList?.contains("is-active") ? "ordered" : "grouped";
    
    // Only apply filter in ordered mode
    if (mode !== "ordered" || !selectedDelimiter) {
      // Show all rows
      const rows = els.editorPreviewTableBody.querySelectorAll("tr");
      for (const row of rows) {
        row.style.display = "";
      }
      return;
    }
    
    // Filter rows based on data-delimiter attribute
    const rows = els.editorPreviewTableBody.querySelectorAll("tr");
    for (const row of rows) {
      const rowDelimiter = row.dataset.delimiter || "";
      if (row.classList.contains("row--delimiter")) {
        // Show delimiter row if it matches
        row.style.display = rowDelimiter === selectedDelimiter ? "" : "none";
      } else {
        // Show data row if it belongs to the selected delimiter
        row.style.display = rowDelimiter === selectedDelimiter ? "" : "none";
      }
    }
  }
  
  function handleDelimiterFilterChanged() {
    applyDelimiterFilter();
  }
  
  function handleEditorColumnSearchChanged() {
    if (!els.editorColumnSearch || !els.editorColumnSearchResults) return;
    const query = String(els.editorColumnSearch.value || "").trim().toLowerCase();
    els.editorColumnSearchResults.innerHTML = "";
    
    if (!query) {
      els.editorColumnSearchResults.style.display = "none";
      return;
    }
    
    // Get existing column options (rows 2-5)
    const opts = Array.isArray(state.editor.columnOptions) ? state.editor.columnOptions : [];
    
    // Also scan row 1 from the workbook
    let row1Options = [];
    try {
      if (state.editor.workbookLoaded && state.workbookArrayBuffer) {
        const wb = ensureWorkbookLoadedForEditor();
        const scope = {
          mode: state.editor.scopeMode,
          sheetName: state.editor.selectedSheetName,
        };
        const scopeMode = scope.mode === "single" ? "single" : "multi";
        const selectedSheet = String(scope.sheetName || "");
        const sheetNames =
          scopeMode === "single"
            ? [selectedSheet].filter(Boolean)
            : wb.SheetNames.slice();
        
        // Scan row 1 for headers
        const row1ByKey = new Map();
        for (const sheetName of sheetNames) {
          const ws = wb.Sheets[sheetName];
          if (!ws) continue;
          const ref = ws["!ref"];
          if (!ref) continue;
          const range = window.XLSX.utils.decode_range(ref);
          const maxCol1 = range.e.c + 1;
          
          // Scan row 1 (row index 0 in 0-based)
          for (let c = 1; c <= maxCol1; c++) {
            const addr = window.XLSX.utils.encode_cell({ r: 0, c: c - 1 });
            const cell = ws[addr];
            if (!cell) continue;
            const raw = cell.v;
            const headerText = String(raw ?? "").trim();
            if (!headerText) continue;
            
            const headerLower = headerText.toLowerCase();
            if (!headerLower.includes(query)) continue;
            
            const key = `unknown::${headerText}`;
            const colLetter = window.XLSX.utils.encode_col(c - 1);
            const loc = { sheet: sheetName, header_row: 1, col1: c, col_letter: colLetter };
            
            if (!row1ByKey.has(key)) {
              row1ByKey.set(key, {
                key,
                headerText,
                kind: "unknown",
                occurrences: 1,
                locations: [loc],
              });
            } else {
              const opt = row1ByKey.get(key);
              opt.occurrences += 1;
              opt.locations.push(loc);
            }
          }
        }
        row1Options = Array.from(row1ByKey.values());
      }
    } catch (e) {
      // If scanning row 1 fails, just use existing options
      console.warn("Failed to scan row 1:", e);
    }
    
    // Combine results from rows 2-5 and row 1
    // Use a map keyed by header text to properly merge
    const allOptionsMap = new Map();
    
    // First, add all options from rows 2-5
    for (const opt of opts) {
      allOptionsMap.set(opt.headerText.toLowerCase(), opt);
    }
    
    // Then merge row 1 options
    for (const row1Opt of row1Options) {
      const headerLower = row1Opt.headerText.toLowerCase();
      const existing = allOptionsMap.get(headerLower);
      if (existing) {
        // Merge row 1 locations into existing option (keep existing key which has correct kind)
        existing.locations.push(...row1Opt.locations);
        existing.occurrences += row1Opt.occurrences;
      } else {
        // Add as new option (row 1 only, so key is unknown::HeaderText)
        allOptionsMap.set(headerLower, row1Opt);
      }
    }
    
    const allOptions = Array.from(allOptionsMap.values());
    
    // Filter options based on search query
    const results = [];
    for (const opt of allOptions) {
      const headerText = String(opt.headerText || "").toLowerCase();
      if (headerText.includes(query)) {
        results.push(opt);
      }
    }
    
    if (results.length === 0) {
      els.editorColumnSearchResults.innerHTML = '<div class="columnSearchResults__empty">No columns found matching your search.</div>';
      els.editorColumnSearchResults.style.display = "block";
      return;
    }
    
    // Display results
    for (const opt of results) {
      const item = document.createElement("div");
      item.className = "columnSearchResults__item";
      item.dataset.columnKey = opt.key;
      
      const tag = opt.kind === "unknown" ? "Unknown" : opt.kind === "lecture" ? "Lecture" : "Section";
      const locationsText = opt.locations && opt.locations.length > 0
        ? opt.locations.map(l => `${l.sheet} (Row ${l.header_row})`).join(", ")
        : "No locations";
      
      item.innerHTML = `
        <div class="columnSearchResults__main">
          <div class="columnSearchResults__header">${escapeHtml(opt.headerText)}</div>
          <div class="columnSearchResults__meta">${tag} • ${opt.occurrences} occurrence(s) • ${locationsText}</div>
        </div>
      `;
      
      els.editorColumnSearchResults.appendChild(item);
    }
    
    els.editorColumnSearchResults.style.display = "block";
  }
  
  function handleEditorColumnSearchResultClicked(e) {
    const item = e?.target?.closest?.(".columnSearchResults__item");
    if (!item) return;
    
    const columnKey = item.dataset.columnKey;
    if (!columnKey) return;
    
    // Set the selected column
    if (els.editorColumn) {
      els.editorColumn.value = columnKey;
      state.editor.selectedColumnKey = columnKey;
    }
    
    // Clear search
    if (els.editorColumnSearch) {
      els.editorColumnSearch.value = "";
    }
    if (els.editorColumnSearchResults) {
      els.editorColumnSearchResults.innerHTML = "";
      els.editorColumnSearchResults.style.display = "none";
    }
    
    // Update UI
    updateWizardUI();
  }

  // ========== OCR Handlers ==========

  /**
   * Handle OCR image upload
   */
  async function handleOcrImageUpload() {
    const input = els.ocrImageUpload;
    if (!input || !input.files || input.files.length === 0) {
      setOcrStatus("Please select at least one image file.", "error");
      return;
    }

    const files = Array.from(input.files);
    state.ocr.uploadedImages = files;
    state.ocr.currentStep = 1;

    // Show image previews
    const previewContainer = els.ocrImagePreview;
    if (previewContainer) {
      previewContainer.innerHTML = "";
      files.forEach((file, index) => {
        const img = document.createElement("img");
        img.src = URL.createObjectURL(file);
        img.style.maxWidth = "200px";
        img.style.maxHeight = "150px";
        img.style.margin = "8px";
        img.style.border = "1px solid var(--border)";
        img.style.borderRadius = "8px";
        previewContainer.appendChild(img);
      });
    }

    if (els.btnOcrProcess) {
      els.btnOcrProcess.disabled = false;
    }

    setOcrStatus(`Loaded ${files.length} image(s). Click "Process Images with OCR" to begin.`, "ok");
  }

  /**
   * Handle OCR processing
   */
  async function handleOcrProcess() {
    if (!state.ocr.uploadedImages || state.ocr.uploadedImages.length === 0) {
      setOcrStatus("Please upload images first.", "error");
      return;
    }

    if (!window.Tesseract) {
      setOcrStatus("Tesseract.js is not loaded. Please refresh the page.", "error");
      return;
    }

    state.ocr.currentStep = 2;
    updateOcrUI();

    try {
      const results = await processMultipleImages(
        state.ocr.uploadedImages,
        (progress, message) => {
          updateOcrProgress(progress, message);
        }
      );

      state.ocr.processingResults = results;
      state.ocr.currentStep = 3;
      updateOcrUI();
      renderOcrResults(results);

      setOcrStatus(
        `Processing complete! Found ${results.confident.length} confident IDs and ${results.uncertain.length} uncertain IDs.`,
        "ok"
      );
    } catch (error) {
      console.error("OCR processing error:", error);
      setOcrStatus(`Error processing images: ${error.message}`, "error");
      state.ocr.currentStep = 1;
      updateOcrUI();
    }
  }

  /**
   * Handle OCR approve and generate text file
   */
  function handleOcrApprove() {
    if (!state.ocr.processingResults) {
      setOcrStatus("No results to approve. Please process images first.", "error");
      return;
    }

    // Collect all approved IDs (confident + manually approved uncertain)
    const approvedIds = [...state.ocr.processingResults.confident];
    
    // Get manually approved uncertain IDs from UI
    const uncertainList = els.ocrUncertainList;
    if (uncertainList) {
      const approvedCheckboxes = uncertainList.querySelectorAll('input[type="checkbox"]:checked');
      approvedCheckboxes.forEach(checkbox => {
        const index = parseInt(checkbox.dataset.index);
        if (state.ocr.processingResults.uncertain[index]) {
          approvedIds.push(state.ocr.processingResults.uncertain[index]);
        }
      });
    }

    if (approvedIds.length === 0) {
      setOcrStatus("No IDs selected for approval.", "error");
      return;
    }

    // Generate text file
    const textContent = generateTextFile(approvedIds);
    const blob = new Blob([textContent], { type: "text/plain" });
    const filename = `attendance_ids_${new Date().toISOString().split('T')[0]}.txt`;
    
    downloadBlob(blob, filename);

    setOcrStatus(`Generated text file with ${approvedIds.length} IDs. File downloaded!`, "ok");
  }

  /**
   * Handle OCR reset
   */
  function handleOcrReset() {
    state.ocr.uploadedImages = [];
    state.ocr.processingResults = null;
    state.ocr.approvedIds = [];
    state.ocr.currentStep = 1;

    if (els.ocrImageUpload) els.ocrImageUpload.value = "";
    if (els.ocrImagePreview) els.ocrImagePreview.innerHTML = "";
    if (els.btnOcrProcess) els.btnOcrProcess.disabled = true;

    updateOcrUI();
    setOcrStatus("Reset complete. You can upload new images.", "info");
  }

  /**
   * Handle uncertain ID edit
   */
  function handleOcrUncertainEdit(index, newId) {
    if (!state.ocr.processingResults || !state.ocr.processingResults.uncertain[index]) {
      return;
    }

    // Update the ID
    state.ocr.processingResults.uncertain[index].id = newId;
    state.ocr.processingResults.uncertain[index].confidence = 100; // Mark as manually edited

    // Re-render results
    renderOcrResults(state.ocr.processingResults);
  }

  /**
   * Update OCR UI based on current step
   */
  function updateOcrUI() {
    // Show/hide steps
    for (let i = 1; i <= 3; i++) {
      const stepEl = document.querySelector(`[data-ocr-step="${i}"]`);
      if (stepEl) {
        stepEl.style.display = state.ocr.currentStep === i ? "block" : "none";
      }
    }
  }

  /**
   * Update OCR progress
   */
  function updateOcrProgress(progress, message) {
    const progressFill = els.ocrProgressFill;
    const progressText = els.ocrProgressText;
    const processingLog = els.ocrProcessingLog;

    if (progressFill) {
      progressFill.style.width = `${Math.min(100, Math.max(0, progress))}%`;
    }
    if (progressText) {
      progressText.textContent = message || `Processing... ${Math.round(progress)}%`;
    }
    if (processingLog) {
      const logEntry = document.createElement("div");
      logEntry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
      logEntry.style.fontSize = "12px";
      logEntry.style.color = "var(--text-secondary)";
      logEntry.style.marginTop = "4px";
      processingLog.appendChild(logEntry);
      processingLog.scrollTop = processingLog.scrollHeight;
    }
  }

  /**
   * Render OCR results
   */
  function renderOcrResults(results) {
    // Update summary
    if (els.ocrTotalIds) {
      els.ocrTotalIds.textContent = results.confident.length + results.uncertain.length;
    }
    if (els.ocrConfidentIds) {
      els.ocrConfidentIds.textContent = results.confident.length;
    }
    if (els.ocrUncertainIds) {
      els.ocrUncertainIds.textContent = results.uncertain.length;
    }

    // Render confident IDs
    const confidentList = els.ocrConfidentList;
    if (confidentList) {
      confidentList.innerHTML = "";
      if (results.confident.length === 0) {
        confidentList.innerHTML = '<p style="color: var(--text-muted);">No confident matches found.</p>';
      } else {
        results.confident.forEach((result, index) => {
          const item = document.createElement("div");
          item.className = "ocrIdItem";
          item.innerHTML = `
            <span class="ocrIdItem__id">${result.id}</span>
            <span class="ocrIdItem__confidence">${Math.round(result.confidence)}%</span>
            <span class="ocrIdItem__source">${result.imageName}</span>
          `;
          confidentList.appendChild(item);
        });
      }
    }

    // Render uncertain IDs with edit capability
    const uncertainList = els.ocrUncertainList;
    if (uncertainList) {
      uncertainList.innerHTML = "";
      if (results.uncertain.length === 0) {
        uncertainList.innerHTML = '<p style="color: var(--text-muted);">No uncertain matches found.</p>';
      } else {
        results.uncertain.forEach((result, index) => {
          const item = document.createElement("div");
          item.className = "ocrIdItem ocrIdItem--uncertain";
          item.innerHTML = `
            <label class="ocrIdItem__checkbox">
              <input type="checkbox" data-index="${index}" />
              <span>Approve</span>
            </label>
            <input 
              type="text" 
              class="ocrIdItem__input" 
              value="${result.id}" 
              data-index="${index}"
              placeholder="Edit ID..."
            />
            <span class="ocrIdItem__confidence">${Math.round(result.confidence)}%</span>
            <span class="ocrIdItem__source">${result.imageName}</span>
            <button class="ocrIdItem__edit" data-index="${index}" type="button">Save</button>
          `;
          uncertainList.appendChild(item);

          // Wire up edit button
          const editBtn = item.querySelector('.ocrIdItem__edit');
          const input = item.querySelector('.ocrIdItem__input');
          if (editBtn && input) {
            editBtn.addEventListener('click', () => {
              handleOcrUncertainEdit(index, input.value);
            });
            input.addEventListener('keypress', (e) => {
              if (e.key === 'Enter') {
                handleOcrUncertainEdit(index, input.value);
              }
            });
          }
        });
      }
    }
  }

  /**
   * Set OCR status message
   */
  function setOcrStatus(msg, kind = "info") {
    const statusEl = els.ocrStatus;
    if (!statusEl) return;
    
    statusEl.textContent = msg || "";
    statusEl.classList.remove("is-error", "is-ok");
    if (kind === "error") statusEl.classList.add("is-error");
    if (kind === "ok") statusEl.classList.add("is-ok");
  }

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
}


