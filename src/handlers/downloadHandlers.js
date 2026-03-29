import {
  FileError,
  ProcessingError,
  ValidationError,
  applyEditorEdits,
} from "../../attendance.js";
import { safeBaseName } from "../metadata.js";
import { downloadBlob } from "../dom.js";
import { readFileAsText } from "../fileRead.js";

/**
 * @param {{
 *   els: any,
 *   state: import('../state.js').AppState,
 *   setEditorStatus: (msg: string, kind?: 'info'|'ok'|'error') => void,
 *   ensureWorkbookLoadedForEditor: () => import('xlsx').WorkBook,
 *   renderEditorPreview: () => void,
 *   switchView: (viewName: string) => void,
 * }} refs
 */
export function createDownloadHandlers(refs) {
  const { els, state, setEditorStatus, ensureWorkbookLoadedForEditor, renderEditorPreview, switchView } = refs;

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
        cellStyles: true,
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
          // Build a map of index -> preview row for quick lookup
          // Use index instead of student_id to handle duplicate IDs correctly
          const rowMap = new Map();
          for (const r of rows) {
            if (r.index !== undefined && r.index !== null && !r.discarded) {
              rowMap.set(Number(r.index), r);
            }
          }

          // Iterate through ordered entries and output modified data
          // Track ID position counter (only increments for ID entries, not titles)
          let idCounter = 0;
          for (const entry of ed.orderedEntries) {
            if (entry && typeof entry === "object") {
              if (entry.type === "title") {
                lines.push(entry.title || "");
              } else if (entry.type === "id") {
                idCounter += 1;
                const previewRow = rowMap.get(idCounter);
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
          // Build a map of index -> preview row for quick lookup
          // Use index instead of input_id/student_id to handle duplicate IDs correctly
          const rowMap = new Map();
          for (const r of rows) {
            if (r.index !== undefined && r.index !== null && !r.discarded) {
              rowMap.set(Number(r.index), r);
            }
          }

          // Iterate through ordered entries and output modified data
          // Track ID position counter (only increments for ID entries, not titles)
          let idCounter = 0;
          for (const entry of ed.orderedEntries) {
            if (entry && typeof entry === "object") {
              if (entry.type === "title") {
                lines.push(entry.title || "");
              } else if (entry.type === "id") {
                idCounter += 1;
                const previewRow = rowMap.get(idCounter);
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
    handleEditorDownloadModified,
    handleDownloadJson,
    handleDownloadTxt,
    handleLoadPreviousReportJson,
    handleDownloadPdf,
    handleDownloadModifiedRecords,
    handleDownloadOriginalRecords,
  };
}
