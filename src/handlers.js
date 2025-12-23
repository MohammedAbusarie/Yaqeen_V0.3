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
} from "../attendance.js";
import { safeBaseName } from "./metadata.js";
import { downloadBlob } from "./dom.js";
import { readFileAsArrayBuffer, readFileAsText } from "./fileRead.js";
import { processMultipleImages, generateTextFile } from "./ocr.js";

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
      if (els.wizardSummaryInput) els.wizardSummaryInput.textContent = ed.inputFileName || "-";
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
        return Boolean(ed.inputFileName);
      case 4:
        return ed.workbookLoaded && Boolean(ed.selectedColumnKey) && Boolean(ed.inputFileName);
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

  function handleEditorInputChanged() {
    const f = els.editorInputTxt?.files?.[0] || null;
    state.editor.inputFileName = f ? f.name : "";
    syncEditorUiFromState();
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

      const inputFile = els.editorInputTxt?.files?.[0] || null;
      if (!inputFile) throw new ValidationError("Input .txt file is required.");

      const inputText = await readFileAsText(inputFile);
      const task = String(ed.taskType || "attendance");

      let preview;
      let orderedEntries = null;
      if (task === "attendance") {
        const parsed = parseStudentIdsText(inputText);
        orderedEntries = parsed.orderedEntries; // Store for delimiter rendering
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
      } else {
        const parsed = parseGradesText(inputText);
        orderedEntries = parsed.orderedEntries; // Store for delimiter rendering
        preview = computeEditorPreview({
          workbook: wb,
          scope,
          columnKey: colKey,
          taskType: "grade",
          orderedAttendanceIds: null,
          attendanceIdsSet: null,
          gradesRows: parsed.rows, // Use rows array from parsed result
        });
      }

      ed.previewRows = preview.preview_rows;
      ed.columnMap = preview.column_map;
      ed.selectedColumn = preview.selected_column;
      ed.orderedEntries = orderedEntries; // Store delimiter information (works for both attendance and grades)
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
      
      // Apply edits with highlight settings
      applyEditorEdits(wb, rows, highlightEnabled, highlightColor);

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
      setEditorStatus(`Downloaded modified file${highlightMsg}. Final column mapping report is shown below.`, "ok");
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

  return {
    handleDownloadJson,
    handleDownloadTxt,
    handleLoadPreviousReportJson,
    handleDownloadPdf,

    // Editor handlers
    handleEditorXlsxUploadChanged,
    handleEditorLoadFile,
    handleEditorSelectionChanged,
    handleEditorTaskChanged,
    handleEditorInputChanged,
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
}


