import {
  DownloadError,
  FileError,
  ProcessingError,
  ValidationError,
  buildStudentSearchRows,
  fetchXlsxFromUrl,
  listColumnOptions,
  parseGradesText,
  parseStudentIdsText,
  computeEditorPreview,
  readWorkbookFromArrayBuffer,
  normalizeId,
} from "../attendance.js";
import { readFileAsArrayBuffer, readFileAsText } from "./fileRead.js";
import { createOcrHandlers } from "./handlers/ocrHandlers.js";
import { createMergerHandlers } from "./handlers/mergerHandlers.js";
import { createDownloadHandlers } from "./handlers/downloadHandlers.js";

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

  const ocr = createOcrHandlers({ els, state });
  const merger = createMergerHandlers({ els, state });
  const downloads = createDownloadHandlers({
    els,
    state,
    setEditorStatus,
    ensureWorkbookLoadedForEditor,
    renderEditorPreview,
    switchView,
  });

  function setEditorStatus(msg, kind = "info") {
    if (!els.editorStatus) return;
    els.editorStatus.textContent = msg || "";
    els.editorStatus.classList.toggle("is-error", kind === "error");
    els.editorStatus.classList.toggle("is-ok", kind === "ok");
  }

  /**
   * Show a short-lived toast message (e.g. "Added: ID — Name"). Does not block interaction.
   * @param {string} message
   * @param {{ duration?: number, kind?: 'ok'|'info'|'error' }} [options]
   */
  function showToast(message, options = {}) {
    const container = els.toastContainer;
    if (!container) return;
    const duration = options.duration ?? 2500;
    const kind = options.kind ?? "ok";
    container.innerHTML = "";
    const el = document.createElement("div");
    el.className = `toast toast--${kind}`;
    el.textContent = message;
    container.appendChild(el);
    setTimeout(() => el.remove(), duration);
  }

  /**
   * Escape text for spreadsheet string literal (double quotes).
   * @param {string} s
   * @returns {string}
   */
  function escapeFormulaString(s) {
    return String(s ?? "").replaceAll('"', '""');
  }

  /**
   * 1-based max row for formula range: preview edits + sheet used range + any cell in target column.
   * Keeps formulas covering the full target column extent (not only rows touched by the preview).
   * @param {import('xlsx').WorkSheet|null|undefined} ws
   * @param {string} targetColLetter
   * @param {number} previewMaxRow
   * @returns {number}
   */
  function getFormulaMaxRowForSheet(ws, targetColLetter, previewMaxRow) {
    const XLSX = window.XLSX;
    let max = Math.max(1, Number(previewMaxRow) || 1);
    if (!ws || !XLSX?.utils) return max;

    if (ws["!ref"]) {
      try {
        const rng = XLSX.utils.decode_range(ws["!ref"]);
        max = Math.max(max, rng.e.r + 1);
      } catch {
        /* ignore */
      }
    }

    const colWant = String(targetColLetter || "A").replace(/\d/g, "").toUpperCase();
    if (!colWant) return Math.min(max, 1048576);

    for (const addr of Object.keys(ws)) {
      if (addr[0] === "!") continue;
      let cell;
      try {
        cell = XLSX.utils.decode_cell(addr);
      } catch {
        continue;
      }
      const letter = XLSX.utils.encode_col(cell.c).toUpperCase();
      if (letter !== colWant) continue;
      max = Math.max(max, cell.r + 1);
    }

    return Math.min(Math.max(1, max), 1048576);
  }

  /**
   * Build per-sheet formula payloads from preview rows.
   * Includes only active rows that are matched/manually fixed and numeric values.
   * @returns {Array<{sheet:string,rowCount:number,skippedCount:number,googleFormula:string,excelFormula:string}>}
   */
  function buildOnlineSheetFormulas() {
    const ed = state.editor || {};
    const rows = Array.isArray(ed.previewRows) ? ed.previewRows : [];
    const headerName = escapeFormulaString(ed.selectedColumn?.headerText || "Target");
    const bySheet = new Map();
    const targetColBySheet = new Map();
    const colMap = Array.isArray(ed.columnMap) ? ed.columnMap : [];
    for (const loc of colMap) {
      const sheet = String(loc?.sheet || "").trim();
      const col = String(loc?.col_letter || "").trim().toUpperCase();
      if (sheet && col) targetColBySheet.set(sheet, col);
    }

    for (const row of rows) {
      const status = String(row?.match_status || "");
      if (row?.discarded) continue;
      if (!(status === "matched" || status === "manuallyFixed")) continue;
      const sheet = String(row?.sheet || "").trim();
      const rowIndex = Number(row?.row_index1);
      const value = Number(row?.new_value);
      if (!sheet || !Number.isFinite(rowIndex) || rowIndex < 1 || !Number.isFinite(value)) {
        const existing = bySheet.get(sheet || "(Unknown sheet)") || { pairs: new Map(), skipped: 0 };
        existing.skipped += 1;
        bySheet.set(sheet || "(Unknown sheet)", existing);
        continue;
      }
      const existing = bySheet.get(sheet) || { pairs: new Map(), skipped: 0 };
      existing.pairs.set(rowIndex, value);
      bySheet.set(sheet, existing);
    }

    const sheets = Array.from(bySheet.keys()).filter((s) => s && s !== "(Unknown sheet)").sort((a, b) => a.localeCompare(b));
    /** @type {import('xlsx').WorkBook|null} */
    let wbForFormulas = null;
    try {
      wbForFormulas = ensureWorkbookLoadedForEditor();
    } catch {
      wbForFormulas = null;
    }

    const out = [];
    for (const sheet of sheets) {
      const info = bySheet.get(sheet);
      const pairs = Array.from(info.pairs.entries()).sort((a, b) => a[0] - b[0]);
      if (!pairs.length) continue;
      const targetCol = targetColBySheet.get(sheet) || "A";
      const previewMaxRow = Math.max(1, ...pairs.map(([r]) => Number(r) || 1));
      const ws = wbForFormulas?.Sheets?.[sheet];
      const maxRow = ws ? getFormulaMaxRowForSheet(ws, targetCol, previewMaxRow) : previewMaxRow;
      const mapLiteral = pairs.map(([r, v]) => `${r},${v}`).join(";");
      const rangeRef = `${targetCol}1:${targetCol}${maxRow}`;
      // Google: IFERROR(VLOOKUP, INDEX) — non-target rows copy existing target column; no LEN (breaks numbers / some values).
      const googleFormula = `=IF(ROW()>${maxRow},"",ARRAYFORMULA(IF(ROW(INDIRECT(ROW()&":"&${maxRow}))=ROW(),"${headerName}",IFERROR(VLOOKUP(ROW(INDIRECT(ROW()&":"&${maxRow})),{${mapLiteral}},2,FALSE),INDEX(${rangeRef},ROW(INDIRECT(ROW()&":"&${maxRow})))))))`;
      // Excel: if no map hit, always use bounded INDEX fallback (full column INDEX + dynamic arrays can fail to copy existing cells).
      const excelFormula = `=LET(_start,ROW(),_n,MAX(1,${maxRow}-_start+1),_rows,SEQUENCE(_n,,_start,1),_m,{${mapLiteral}},_hit,IFERROR(XLOOKUP(_rows,INDEX(_m,,1),INDEX(_m,,2)),""),_fallback,INDEX(${rangeRef},_rows),IF(_rows=_start,"${headerName}",IF(_hit<>"",_hit,_fallback)))`;
      out.push({
        sheet,
        targetCol,
        maxRow,
        previewMaxRow,
        rowCount: pairs.length,
        skippedCount: info.skipped,
        googleFormula,
        excelFormula,
      });
    }
    return out;
  }

  function renderOnlineSheetFormulaPanel() {
    if (!els.editorFormulaPanels) return;
    const formulas = buildOnlineSheetFormulas();
    els.editorFormulaPanels.innerHTML = "";
    if (!formulas.length) {
      const empty = document.createElement("div");
      empty.className = "formulaPanel__empty";
      empty.textContent =
        "No eligible rows for formulas yet. Eligible rows are matched/manually fixed and not discarded, with numeric values.";
      els.editorFormulaPanels.appendChild(empty);
      return;
    }

    for (const item of formulas) {
      const card = document.createElement("article");
      card.className = "formulaCard";

      const title = document.createElement("div");
      title.className = "formulaCard__title";
      title.textContent = `${item.sheet} (${item.rowCount} target rows)`;
      card.appendChild(title);

      const targetNote = document.createElement("div");
      targetNote.className = "formulaCard__note";
      const ext =
        item.maxRow > item.previewMaxRow
          ? ` Spans full sheet/target column through row ${item.maxRow} (preview edits only reached row ${item.previewMaxRow}).`
          : ` Covers rows 1–${item.maxRow}.`;
      targetNote.textContent = `Row-anchored; range ${item.targetCol}1:${item.targetCol}${item.maxRow}.${ext} Non-target rows copy existing values (IFERROR map miss → INDEX). If you loaded preview from JSON only, range may be limited to preview rows—reload the workbook to extend.`;
      card.appendChild(targetNote);

      if (item.skippedCount > 0) {
        const note = document.createElement("div");
        note.className = "formulaCard__note";
        note.textContent = `${item.skippedCount} row(s) were skipped because their new value is not numeric.`;
        card.appendChild(note);
      }

      const gLabel = document.createElement("div");
      gLabel.className = "formulaCard__label";
      gLabel.textContent = "Google Sheets formula";
      card.appendChild(gLabel);

      const gArea = document.createElement("textarea");
      gArea.className = "formulaCard__textarea";
      gArea.readOnly = true;
      gArea.rows = 3;
      gArea.value = item.googleFormula;
      card.appendChild(gArea);

      const gBtn = document.createElement("button");
      gBtn.type = "button";
      gBtn.className = "btn btn--ghost formulaCard__copyBtn";
      gBtn.dataset.copyValue = item.googleFormula;
      gBtn.textContent = "Copy Google formula";
      card.appendChild(gBtn);

      const eLabel = document.createElement("div");
      eLabel.className = "formulaCard__label";
      eLabel.textContent = "Excel formula";
      card.appendChild(eLabel);

      const eArea = document.createElement("textarea");
      eArea.className = "formulaCard__textarea";
      eArea.readOnly = true;
      eArea.rows = 3;
      eArea.value = item.excelFormula;
      card.appendChild(eArea);

      const eBtn = document.createElement("button");
      eBtn.type = "button";
      eBtn.className = "btn btn--ghost formulaCard__copyBtn";
      eBtn.dataset.copyValue = item.excelFormula;
      eBtn.textContent = "Copy Excel formula";
      card.appendChild(eBtn);

      els.editorFormulaPanels.appendChild(card);
    }
  }

  /** Column option value encoding: key + location so dropdown and search stay in sync. Delimiter must not appear in sheet names. */
  const COLUMN_VALUE_SEP = "||";

  /**
   * Deduplicate locations by (sheet, header_row, col_letter).
   * @param {Array<{ sheet: string, header_row: number, col_letter: string }>} locations
   * @returns {typeof locations}
   */
  function deduplicateLocations(locations) {
    if (!Array.isArray(locations) || locations.length === 0) return locations;
    const seen = new Set();
    return locations.filter((loc) => {
      const key = `${loc.sheet}\t${loc.header_row}\t${loc.col_letter}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
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

    // input method radios and container visibility
    const method = ed.inputMethod || "file";
    if (els.editorInputMethodFile) els.editorInputMethodFile.checked = method === "file";
    if (els.editorInputMethodTextarea) els.editorInputMethodTextarea.checked = method === "textarea";
    if (els.editorInputMethodSearchPick) els.editorInputMethodSearchPick.checked = method === "searchPick";
    if (els.editorInputFileContainer) els.editorInputFileContainer.style.display = method === "file" ? "block" : "none";
    if (els.editorInputTextareaContainer) els.editorInputTextareaContainer.style.display = method === "textarea" ? "block" : "none";
    if (els.editorInputSearchPickContainer) els.editorInputSearchPickContainer.style.display = method === "searchPick" ? "block" : "none";

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
        if (ed.inputMethod === "file") {
          els.wizardSummaryInput.textContent = ed.inputFileName || "-";
        } else if (ed.inputMethod === "searchPick") {
          const n = Array.isArray(ed.chosenStudents) ? ed.chosenStudents.length : 0;
          els.wizardSummaryInput.textContent = n > 0 ? `Search & pick: ${n} students` : "-";
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
        if (ed.inputMethod === "file") {
          return Boolean(els.editorInputTxt?.files?.[0]);
        }
        if (ed.inputMethod === "searchPick") {
          return Array.isArray(ed.chosenStudents) && ed.chosenStudents.length > 0;
        }
        return Boolean(ed.inputTextContent && ed.inputTextContent.trim());
      case 4:
        return ed.workbookLoaded && Boolean(ed.selectedColumnKey) && (
          (ed.inputMethod === "file" && Boolean(els.editorInputTxt?.files?.[0])) ||
          (ed.inputMethod === "textarea" && Boolean(ed.inputTextContent && ed.inputTextContent.trim())) ||
          (ed.inputMethod === "searchPick" && Array.isArray(ed.chosenStudents) && ed.chosenStudents.length > 0)
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
    renderOnlineSheetFormulaPanel();

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

  async function handleEditorFormulaCopyClicked(e) {
    const btn = e?.target?.closest?.("button[data-copy-value]");
    if (!btn) return;
    const text = String(btn.dataset.copyValue || "");
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      showToast("Formula copied.");
    } catch {
      showToast("Copy failed. Select the formula and copy manually.", { kind: "error" });
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
          const tag = opt.kind === "unknown" ? "Unknown" : opt.kind === "lecture" ? "Lecture" : "Section";
          const locs = deduplicateLocations(Array.isArray(opt.locations) ? opt.locations : []);
          for (const loc of locs) {
            const composite = `${opt.key}${COLUMN_VALUE_SEP}${loc.sheet}${COLUMN_VALUE_SEP}${loc.header_row}${COLUMN_VALUE_SEP}${loc.col_letter}`;
            const o = document.createElement("option");
            o.value = composite;
            o.textContent = `${opt.headerText} — ${tag} — ${loc.sheet} (Row ${loc.header_row}, Col ${loc.col_letter})`;
            els.editorColumn.appendChild(o);
          }
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
      editorSearchRows = [];
      state.editor.chosenStudents = [];
      state.editor.selectedLocation = null;
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
        const rawValue = String(els.editorColumn.value || "").trim();
        const parts = rawValue ? rawValue.split(COLUMN_VALUE_SEP) : [];
        if (parts.length === 4) {
          state.editor.selectedColumnKey = parts[0];
          state.editor.selectedLocation = {
            sheet: parts[1],
            header_row: Number.parseInt(parts[2], 10),
            col_letter: parts[3],
          };
        } else {
          state.editor.selectedColumnKey = rawValue || "";
          state.editor.selectedLocation = null;
        }
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
        editorSearchRows = [];
        const wb = ensureWorkbookLoadedForEditor();
        const scope = { mode: state.editor.scopeMode, sheetName: state.editor.selectedSheetName };
        const opts = listColumnOptions(wb, scope);

        // Store column options for search functionality
        state.editor.columnOptions = opts;

        if (els.editorColumn) {
          els.editorColumn.innerHTML = "";
          const o0 = document.createElement("option");
          o0.value = "";
          o0.textContent = "(Select column)";
          els.editorColumn.appendChild(o0);
          const ed = state.editor;
          let selectedValue = "";
          for (const opt of opts) {
            const tag = opt.kind === "unknown" ? "Unknown" : opt.kind === "lecture" ? "Lecture" : "Section";
            const locs = deduplicateLocations(Array.isArray(opt.locations) ? opt.locations : []);
            for (const loc of locs) {
              const composite = `${opt.key}${COLUMN_VALUE_SEP}${loc.sheet}${COLUMN_VALUE_SEP}${loc.header_row}${COLUMN_VALUE_SEP}${loc.col_letter}`;
              const o = document.createElement("option");
              o.value = composite;
              o.textContent = `${opt.headerText} — ${tag} — ${loc.sheet} (Row ${loc.header_row}, Col ${loc.col_letter})`;
              els.editorColumn.appendChild(o);
              if (ed.selectedLocation && ed.selectedColumnKey === opt.key &&
                  ed.selectedLocation.sheet === loc.sheet &&
                  ed.selectedLocation.header_row === loc.header_row &&
                  ed.selectedLocation.col_letter === loc.col_letter) {
                selectedValue = composite;
              }
              if (!selectedValue && ed.selectedColumnKey === opt.key && !ed.selectedLocation) {
                selectedValue = composite; // first occurrence for this key
              }
            }
          }
          if (selectedValue && Array.from(els.editorColumn.options).some((opt) => opt.value === selectedValue)) {
            els.editorColumn.value = selectedValue;
          } else {
            els.editorColumn.value = "";
            if (!selectedValue && ed.selectedColumnKey) {
              state.editor.selectedColumnKey = "";
              state.editor.selectedLocation = null;
            }
          }
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
    if (state.editor.inputMethod === "searchPick") {
      renderEditorChosenList();
    }
    syncEditorUiFromState();
  }

  function handleEditorInputMethodChanged() {
    const isFile = els.editorInputMethodFile?.checked;
    const isSearchPick = els.editorInputMethodSearchPick?.checked;
    state.editor.inputMethod = isSearchPick ? "searchPick" : (isFile ? "file" : "textarea");

    if (!isSearchPick) {
      state.editor.chosenStudents = [];
    }

    if (els.editorInputFileContainer) {
      els.editorInputFileContainer.style.display = isFile ? "block" : "none";
    }
    if (els.editorInputTextareaContainer) {
      els.editorInputTextareaContainer.style.display = isSearchPick ? "none" : (isFile ? "none" : "block");
    }
    if (els.editorInputSearchPickContainer) {
      els.editorInputSearchPickContainer.style.display = isSearchPick ? "block" : "none";
    }

    if (isSearchPick) {
      ensurePickSearchRows();
      renderEditorChosenList();
      if (els.editorPickSearch) els.editorPickSearch.value = "";
      if (els.editorPickResults) els.editorPickResults.innerHTML = "";
    }

    syncEditorUiFromState();
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

  function ensurePickSearchRows() {
    if (editorSearchRows.length > 0) return;
    try {
      const wb = ensureWorkbookLoadedForEditor();
      const ed = state.editor;
      const scope = { mode: ed.scopeMode, sheetName: ed.selectedSheetName };
      editorSearchRows = buildStudentSearchRows(wb, scope);
    } catch (e) {
      editorSearchRows = [];
    }
  }

  function renderEditorChosenList() {
    const container = els.editorChosenList;
    const emptyEl = els.editorChosenListEmpty;
    if (!container) return;
    const chosen = state.editor.chosenStudents || [];
    const taskType = state.editor.taskType || "attendance";

    if (emptyEl) {
      emptyEl.style.display = chosen.length === 0 ? "block" : "none";
    }

    container.innerHTML = "";
    chosen.forEach((s, index) => {
      const row = document.createElement("div");
      row.className = "chosenList__item";
      row.dataset.index = String(index);
      const id = String(s.id || "");
      const name = String(s.name || "").trim() || "—";
      const grade = taskType === "grade" ? String(s.grade ?? "").trim() : "";
      const text = taskType === "grade" && grade
        ? `${id} — ${name} — ${grade}`
        : `${id} — ${name}`;
      const span = document.createElement("span");
      span.className = "chosenList__text";
      span.textContent = text;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn btn--ghost chosenList__remove";
      btn.textContent = "Remove";
      btn.dataset.index = String(index);
      btn.setAttribute("aria-label", `Remove ${id} from list`);
      row.appendChild(span);
      row.appendChild(btn);
      container.appendChild(row);
    });
  }

  function handleEditorPickSearchChanged() {
    if (!els.editorPickSearch || !els.editorPickResults) return;
    const q = String(els.editorPickSearch.value || "").trim().toLowerCase();
    els.editorPickResults.innerHTML = "";
    if (!q) return;

    ensurePickSearchRows();
    if (editorSearchRows.length === 0) {
      els.editorPickResults.innerHTML = '<div class="searchResults__empty">No student data. Load workbook and select sheet/column first.</div>';
      return;
    }

    const results = [];
    for (const r of editorSearchRows) {
      const id = String(r.id || "");
      const name = String(r.name || "");
      if (id.toLowerCase().includes(q) || name.toLowerCase().includes(q)) results.push(r);
      if (results.length >= 30) break;
    }

    if (results.length === 0) {
      els.editorPickResults.innerHTML = '<div class="searchResults__empty">No students found matching your search.</div>';
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
      els.editorPickResults.appendChild(item);
    }
  }

  let editorPickGradePending = null;

  function resetEditorPickSearch({ clearValue = false, clearResults = false } = {}) {
    if (els.editorPickSearch) {
      if (clearValue) els.editorPickSearch.value = "";
      els.editorPickSearch.focus();
    }
    if (clearResults && els.editorPickResults) {
      els.editorPickResults.innerHTML = "";
    }
  }

  function handleEditorPickResultClicked(e) {
    const node = e?.target?.closest?.(".searchResults__item");
    if (!node || !node.closest("#editorPickResults")) return;
    const sheet = String(node.dataset.sheet || "");
    const row1 = Number.parseInt(String(node.dataset.row1 || ""), 10);
    const id = String(node.dataset.id || "");
    const name = String(node.dataset.name || "");

    const chosen = state.editor.chosenStudents || [];
    const isDuplicate = chosen.some((s) => String(s.id) === id);
    if (isDuplicate) {
      showToast("Already in list: " + id + (name ? " — " + name : ""), { kind: "info" });
      resetEditorPickSearch({ clearValue: false, clearResults: false });
      return;
    }

    const taskType = state.editor.taskType || "attendance";
    if (taskType === "attendance") {
      state.editor.chosenStudents = [...chosen, { id, name, sheet, row1 }];
      renderEditorChosenList();
      updateWizardUI();
      showToast("Added: " + id + (name ? " — " + name : ""));
      resetEditorPickSearch({ clearValue: true, clearResults: true });
      return;
    }

    editorPickGradePending = { id, name, sheet, row1 };
    if (els.editorPickGradeDialogTitle) {
      els.editorPickGradeDialogTitle.textContent = `Grade for ${id} — ${name}`;
    }
    if (els.editorPickGradeValue) {
      els.editorPickGradeValue.value = "";
    }
    els.editorPickGradeDialog?.showModal();
    els.editorPickGradeValue?.focus();
  }

  function handleEditorPickGradeAdd() {
    if (!editorPickGradePending) return;
    const grade = String(els.editorPickGradeValue?.value ?? "").trim();
    if (!grade) {
      setEditorStatus("Please enter a grade.", "error");
      return;
    }
    const chosen = state.editor.chosenStudents || [];
    const id = String(editorPickGradePending.id || "");
    const name = String(editorPickGradePending.name || "");
    if (chosen.some((s) => String(s.id) === id)) {
      editorPickGradePending = null;
      els.editorPickGradeDialog?.close();
      showToast("Already in list: " + id + (name ? " — " + name : ""), { kind: "info" });
      resetEditorPickSearch({ clearValue: false, clearResults: false });
      return;
    }
    state.editor.chosenStudents = [...chosen, { ...editorPickGradePending, grade }];
    editorPickGradePending = null;
    els.editorPickGradeDialog?.close();
    renderEditorChosenList();
    updateWizardUI();
    setEditorStatus("");
    showToast("Added: " + id + (name ? " — " + name : ""));
    resetEditorPickSearch({ clearValue: true, clearResults: true });
  }

  function handleEditorChosenListRemove(e) {
    const btn = e?.target?.closest?.(".chosenList__remove");
    if (!btn) return;
    const index = Number.parseInt(String(btn.dataset.index || ""), 10);
    if (!Number.isFinite(index)) return;
    const chosen = [...(state.editor.chosenStudents || [])];
    if (index < 0 || index >= chosen.length) return;
    chosen.splice(index, 1);
    state.editor.chosenStudents = chosen;
    renderEditorChosenList();
    updateWizardUI();
  }

  function handleEditorPickGradeDialogClosed() {
    editorPickGradePending = null;
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

      const task = String(ed.taskType || "attendance");
      let preview;
      let orderedEntries = null;
      let idCounts = null;

      if (ed.inputMethod === "searchPick") {
        const chosen = ed.chosenStudents || [];
        if (chosen.length === 0) throw new ValidationError("Add at least one student from the search results.");

        if (task === "attendance") {
          orderedEntries = chosen.map((s) => ({ type: "id", id: String(s.id) }));
          const targetIdsSet = new Set(chosen.map((s) => String(s.id)));
          idCounts = {};
          for (const s of chosen) {
            const sid = normalizeId(s.id);
            if (sid) idCounts[sid] = (idCounts[sid] || 0) + 1;
          }
          const orderedIds = chosen.map((s) => String(s.id));
          preview = computeEditorPreview({
            workbook: wb,
            scope,
            columnKey: colKey,
            selectedLocation: ed.selectedLocation || undefined,
            taskType: "attendance",
            orderedAttendanceIds: orderedIds,
            attendanceIdsSet: targetIdsSet,
            gradesRows: null,
          });
          ed.originalInputData = { type: "attendance", orderedEntries, idsSet: targetIdsSet };
        } else {
          const missingGrade = chosen.find((s) => !String(s.grade ?? "").trim());
          if (missingGrade) {
            throw new ValidationError("Enter a grade for each chosen student.");
          }
          const rows = chosen.map((s) => ({ id: String(s.id), grade: String(s.grade ?? "").trim() }));
          orderedEntries = rows.map((r) => ({ type: "id", id: r.id, grade: r.grade }));
          idCounts = {};
          for (const s of chosen) {
            const sid = normalizeId(s.id);
            if (sid) idCounts[sid] = (idCounts[sid] || 0) + 1;
          }
          preview = computeEditorPreview({
            workbook: wb,
            scope,
            columnKey: colKey,
            selectedLocation: ed.selectedLocation || undefined,
            taskType: "grade",
            orderedAttendanceIds: null,
            attendanceIdsSet: null,
            gradesRows: rows,
          });
          ed.originalInputData = { type: "grade", orderedEntries, rows };
        }
        ed.previewRows = preview.preview_rows;
        ed.columnMap = preview.column_map;
        ed.selectedColumn = preview.selected_column;
        ed.orderedEntries = orderedEntries;
        ed.idCounts = idCounts;
        editorSearchRows = buildStudentSearchRows(wb, scope);
        if (els.editorPreviewSheetFilter) {
          els.editorPreviewSheetFilter.disabled = false;
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
        if (els.btnEditorDownload) els.btnEditorDownload.disabled = false;
        if (els.btnDownloadModifiedRecords) els.btnDownloadModifiedRecords.disabled = false;
        if (els.btnDownloadOriginalRecords) els.btnDownloadOriginalRecords.disabled = false;
        if (els.btnDownloadJson) els.btnDownloadJson.disabled = false;
        if (els.btnDownloadTxt) els.btnDownloadTxt.disabled = false;
        if (els.btnDownloadPdf) els.btnDownloadPdf.disabled = false;
        renderEditorPreview();
        syncEditorUiFromState();
        setEditorStatus("Preview generated. Review carefully, then download when ready.", "ok");
        switchView("report");
        return;
      }

      // Get input text - either from file or textarea
      let inputText;
      if (ed.inputMethod === "file") {
        const inputFile = els.editorInputTxt?.files?.[0] || null;
        if (!inputFile) throw new ValidationError("Input .txt file is required.");
        inputText = await readFileAsText(inputFile);
      } else {
        inputText = ed.inputTextContent;
        if (!inputText || !inputText.trim()) {
          throw new ValidationError("Please enter input data in the text area.");
        }
      }

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
          selectedLocation: ed.selectedLocation || undefined,
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
          selectedLocation: ed.selectedLocation || undefined,
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

  return {
    handleDownloadJson: downloads.handleDownloadJson,
    handleDownloadTxt: downloads.handleDownloadTxt,
    handleLoadPreviousReportJson: downloads.handleLoadPreviousReportJson,
    handleDownloadPdf: downloads.handleDownloadPdf,
    handleDownloadModifiedRecords: downloads.handleDownloadModifiedRecords,
    handleDownloadOriginalRecords: downloads.handleDownloadOriginalRecords,

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
    handleEditorDownloadModified: downloads.handleEditorDownloadModified,
    handleEditorFormulaCopyClicked,

    // Preview fixups
    handleEditorPreviewRowAction,
    handleEditorFixSearchChanged,
    handleEditorFixResultClicked,
    handleEditorGradeSaveClicked,
    handleDelimiterFilterChanged,

    // Search & pick
    handleEditorPickSearchChanged,
    handleEditorPickResultClicked,
    handleEditorPickGradeAdd,
    handleEditorChosenListRemove,
    handleEditorPickGradeDialogClosed,

    // Wizard navigation
    handleWizardPrev,
    handleWizardNext,
    handleWizardFinish,
    updateWizardUI,
    
    // Column search
    handleEditorColumnSearchChanged,
    handleEditorColumnSearchResultClicked,
    
    // OCR handlers
    handleOcrImageUpload: ocr.handleOcrImageUpload,
    handleOcrProcess: ocr.handleOcrProcess,
    handleOcrApprove: ocr.handleOcrApprove,
    handleOcrReset: ocr.handleOcrReset,
    updateOcrUI: ocr.updateOcrUI,

    // Sheet Merger handlers
    handleMergerLoadFile: merger.handleMergerLoadFile,
    handleMergerColumnDragStart: merger.handleMergerColumnDragStart,
    handleMergerMatrixDragOver: merger.handleMergerMatrixDragOver,
    handleMergerMatrixDrop: merger.handleMergerMatrixDrop,
    handleMergerMatrixClick: merger.handleMergerMatrixClick,
    handleMergerEliminateHeaders: merger.handleMergerEliminateHeaders,
    handleMergerGeneratePreview: merger.handleMergerGeneratePreview,
    handleMergerDownload: merger.handleMergerDownload,
    handleMergerReset: merger.handleMergerReset,
    handleMergerBack: merger.handleMergerBack,
    handleMergerLoadMore: merger.handleMergerLoadMore,
    handleMergerColumnSearch: merger.handleMergerColumnSearch,
    handleMergerColumnGroupToggle: merger.handleMergerColumnGroupToggle,
    handleMergerAutoFill: merger.handleMergerAutoFill,
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
    
    // Then merge row 1 options (avoid duplicating locations already in listColumnOptions row 1)
    for (const row1Opt of row1Options) {
      const headerLower = row1Opt.headerText.toLowerCase();
      const existing = allOptionsMap.get(headerLower);
      if (existing) {
        const existingSet = new Set(
          (existing.locations || []).map((l) => `${l.sheet}\t${l.header_row}\t${l.col_letter}`)
        );
        for (const loc of row1Opt.locations || []) {
          const key = `${loc.sheet}\t${loc.header_row}\t${loc.col_letter}`;
          if (!existingSet.has(key)) {
            existing.locations.push(loc);
            existingSet.add(key);
          }
        }
        existing.occurrences = existing.locations.length;
      } else {
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

    // Build flat list of occurrences (one row per location), deduplicated so each physical column appears once
    /** @type {Array<{ opt: typeof results[0], loc: { sheet: string, header_row: number, col1: number, col_letter: string } }>} */
    const occurrenceRows = [];
    for (const opt of results) {
      const locations = deduplicateLocations(Array.isArray(opt.locations) ? opt.locations : []);
      locations.forEach((loc) => {
        occurrenceRows.push({ opt, loc });
      });
    }

    for (const { opt, loc } of occurrenceRows) {
      const item = document.createElement("div");
      item.className = "columnSearchResults__item";
      item.dataset.columnKey = opt.key;
      item.dataset.locationSheet = loc.sheet;
      item.dataset.locationHeaderRow = String(loc.header_row);
      item.dataset.locationColLetter = loc.col_letter;

      const tag = opt.kind === "unknown" ? "Unknown" : opt.kind === "lecture" ? "Lecture" : "Section";
      const positionText = `${loc.sheet} (Row ${loc.header_row}, Col ${loc.col_letter})`;

      item.innerHTML = `
        <div class="columnSearchResults__main">
          <div class="columnSearchResults__header">${escapeHtml(opt.headerText)} — ${tag}</div>
          <div class="columnSearchResults__meta">${escapeHtml(positionText)}</div>
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

    const sheet = item.dataset.locationSheet;
    const headerRow = Number.parseInt(item.dataset.locationHeaderRow || "0", 10);
    const colLetter = item.dataset.locationColLetter;
    const hasLocation = sheet && Number.isFinite(headerRow) && colLetter;

    state.editor.selectedColumnKey = columnKey;
    state.editor.selectedLocation = hasLocation
      ? { sheet, header_row: headerRow, col_letter: colLetter }
      : null;

    if (els.editorColumn && hasLocation) {
      const composite = `${columnKey}${COLUMN_VALUE_SEP}${sheet}${COLUMN_VALUE_SEP}${headerRow}${COLUMN_VALUE_SEP}${colLetter}`;
      if (Array.from(els.editorColumn.options).some((opt) => opt.value === composite)) {
        els.editorColumn.value = composite;
      } else {
        els.editorColumn.value = "";
      }
    } else if (els.editorColumn) {
      els.editorColumn.value = "";
    }

    if (els.editorColumnSearch) {
      els.editorColumnSearch.value = "";
    }
    if (els.editorColumnSearchResults) {
      els.editorColumnSearchResults.innerHTML = "";
      els.editorColumnSearchResults.style.display = "none";
    }

    if (hasLocation && setEditorStatus) {
      setEditorStatus(`Using: ${sheet} Row ${headerRow}, Col ${colLetter}`, "ok");
    }

    updateWizardUI();
  }

}


