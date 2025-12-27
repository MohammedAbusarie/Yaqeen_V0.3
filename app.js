import { $ as domGet } from "./src/dom.js";
import { disableRun as disableRunUi, setStatus as setStatusUi } from "./src/uiStatus.js";
import { createInitialState, resetAll as resetAllState } from "./src/state.js";
import { switchView as switchViewNav } from "./src/navigation.js";
import { createHandlers } from "./src/handlers.js";

const els = {
  // Navigation
  navInputs: domGet("navInputs"),
  navReport: domGet("navReport"),
  navOcr: domGet("navOcr"),
  navAbout: domGet("navAbout"),
  viewInputs: domGet("viewInputs"),
  viewReport: domGet("viewReport"),
  viewOcr: domGet("viewOcr"),
  viewAbout: domGet("viewAbout"),
  
  // Editor (generic column edit)
  editorScope: domGet("editorScope"),
  editorSheet: domGet("editorSheet"),
  editorColumn: domGet("editorColumn"),
  editorColumnSearch: domGet("editorColumnSearch"),
  editorColumnSearchResults: domGet("editorColumnSearchResults"),
  editorTask: domGet("editorTask"),
  editorSheetUrl: domGet("editorSheetUrl"),
  editorXlsxFile: domGet("editorXlsxFile"),
  editorInputMethodFile: domGet("editorInputMethodFile"),
  editorInputMethodTextarea: domGet("editorInputMethodTextarea"),
  editorInputFileContainer: domGet("editorInputFileContainer"),
  editorInputTextareaContainer: domGet("editorInputTextareaContainer"),
  editorInputTxt: domGet("editorInputTxt"),
  editorInputTextarea: domGet("editorInputTextarea"),
  btnEditorDownload: domGet("btnEditorDownload"),
  btnDownloadModifiedRecords: domGet("btnDownloadModifiedRecords"),
  btnDownloadOriginalRecords: domGet("btnDownloadOriginalRecords"),
  editorStatus: domGet("editorStatus"),

  // Wizard
  btnWizardPrev: domGet("btnWizardPrev"),
  btnWizardNext: domGet("btnWizardNext"),
  btnWizardFinish: domGet("btnWizardFinish"),
  wizardSummaryFile: domGet("wizardSummaryFile"),
  wizardSummaryMode: domGet("wizardSummaryMode"),
  wizardSummarySheet: domGet("wizardSummarySheet"),
  wizardSummaryColumn: domGet("wizardSummaryColumn"),
  wizardSummaryTask: domGet("wizardSummaryTask"),
  wizardSummaryInput: domGet("wizardSummaryInput"),

  // Preview table
  editorPreviewModeGrouped: domGet("editorPreviewModeGrouped"),
  editorPreviewModeOrdered: domGet("editorPreviewModeOrdered"),
  editorPreviewSheetFilter: domGet("editorPreviewSheetFilter"),
  editorPreviewTableBody: domGet("editorPreviewTableBody"),
  editorFinalReportBox: domGet("editorFinalReportBox"),
  delimiterFilterContainer: domGet("delimiterFilterContainer"),
  editorDelimiterFilter: domGet("editorDelimiterFilter"),

  // Fix dialogs
  editorFixDialog: domGet("editorFixDialog"),
  editorFixSearch: domGet("editorFixSearch"),
  editorFixResults: domGet("editorFixResults"),
  editorGradeDialog: domGet("editorGradeDialog"),
  editorGradeValue: domGet("editorGradeValue"),
  editorGradeSave: domGet("editorGradeSave"),

  // Highlight controls
  highlightCheckbox: domGet("highlightCheckbox"),
  highlightColorPicker: domGet("highlightColorPicker"),

  // Report
  summary: domGet("summary"),
  btnDownloadJson: domGet("btnDownloadJson"),
  btnDownloadTxt: domGet("btnDownloadTxt"),
  btnDownloadPdf: domGet("btnDownloadPdf"),
  loadReportJson: domGet("loadReportJson"),

  // OCR
  ocrImageUpload: domGet("ocrImageUpload"),
  ocrImagePreview: domGet("ocrImagePreview"),
  btnOcrProcess: domGet("btnOcrProcess"),
  ocrProgressFill: domGet("ocrProgressFill"),
  ocrProgressText: domGet("ocrProgressText"),
  ocrProcessingLog: domGet("ocrProcessingLog"),
  ocrTotalIds: domGet("ocrTotalIds"),
  ocrConfidentIds: domGet("ocrConfidentIds"),
  ocrUncertainIds: domGet("ocrUncertainIds"),
  ocrConfidentList: domGet("ocrConfidentList"),
  ocrUncertainList: domGet("ocrUncertainList"),
  btnOcrApprove: domGet("btnOcrApprove"),
  btnOcrReset: domGet("btnOcrReset"),
  ocrStatus: domGet("ocrStatus"),
};

const state = createInitialState();

// View management
function switchView(viewName) {
  switchViewNav(domGet, viewName);
}

function setStatus(msg, kind = "info") {
  setStatusUi(els, msg, kind);
}

function disableRun(disabled, label = null) {
  disableRunUi(els, disabled, label);
}

function resetAll() {
  resetAllState(els, setStatus, state);
  if (handlers.updateWizardUI) handlers.updateWizardUI();
}

const handlers = createHandlers({
  els,
  state,
  setStatus,
  disableRun,
  switchView,
});

// Wire up events
if (els.navInputs) {
  els.navInputs.addEventListener("click", () => switchView('inputs'));
} else {
  console.error("navInputs element not found");
}
if (els.navReport) {
  els.navReport.addEventListener("click", () => switchView('report'));
} else {
  console.error("navReport element not found");
}
if (els.navOcr) {
  els.navOcr.addEventListener("click", () => switchView('ocr'));
} else {
  console.error("navOcr element not found");
}
if (els.navAbout) {
  els.navAbout.addEventListener("click", () => switchView('about'));
} else {
  console.error("navAbout element not found");
}

els.btnDownloadJson?.addEventListener("click", handlers.handleDownloadJson);
els.btnDownloadTxt?.addEventListener("click", handlers.handleDownloadTxt);
els.btnDownloadPdf?.addEventListener("click", handlers.handleDownloadPdf);
els.btnDownloadModifiedRecords?.addEventListener("click", handlers.handleDownloadModifiedRecords);
els.btnDownloadOriginalRecords?.addEventListener("click", handlers.handleDownloadOriginalRecords);
els.loadReportJson?.addEventListener("change", handlers.handleLoadPreviousReportJson);

// Editor wiring (Inputs view)
els.editorSheetUrl?.addEventListener("input", () => {
  // Clear workbook buffer when URL changes (user wants to use URL instead of file)
  if (state.workbookArrayBuffer && !els.editorXlsxFile?.files?.[0]) {
    state.workbookArrayBuffer = null;
    state.workbookName = null;
    state.editor.workbookLoaded = false;
  }
  // Update wizard UI to refresh Next button state
  handlers.updateWizardUI?.();
});
els.editorXlsxFile?.addEventListener("change", handlers.handleEditorXlsxUploadChanged);
// Load button removed - loading happens automatically when Next is pressed on step 1
els.editorScope?.addEventListener("change", handlers.handleEditorSelectionChanged);
els.editorSheet?.addEventListener("change", handlers.handleEditorSelectionChanged);
els.editorColumn?.addEventListener("change", handlers.handleEditorSelectionChanged);
els.editorTask?.addEventListener("change", handlers.handleEditorTaskChanged);
els.editorInputMethodFile?.addEventListener("change", handlers.handleEditorInputMethodChanged);
els.editorInputMethodTextarea?.addEventListener("change", handlers.handleEditorInputMethodChanged);
els.editorInputTxt?.addEventListener("change", handlers.handleEditorInputChanged);
els.editorInputTextarea?.addEventListener("input", handlers.handleEditorTextareaChanged);

// Editor wiring (Reports view - buttons are in Reports view now)
els.btnEditorDownload?.addEventListener("click", handlers.handleEditorDownloadModified);

// Highlight controls
els.highlightCheckbox?.addEventListener("change", (e) => {
  const checked = e.target.checked;
  state.editor.highlightEnabled = checked;
  if (els.highlightColorPicker) {
    els.highlightColorPicker.disabled = !checked;
  }
});

els.highlightColorPicker?.addEventListener("change", (e) => {
  state.editor.highlightColor = e.target.value;
});

els.editorPreviewModeGrouped?.addEventListener("click", handlers.handleEditorPreviewModeChanged);
els.editorPreviewModeOrdered?.addEventListener("click", handlers.handleEditorPreviewModeChanged);
els.editorPreviewSheetFilter?.addEventListener("change", handlers.handleEditorPreviewModeChanged);
els.editorPreviewTableBody?.addEventListener("click", handlers.handleEditorPreviewRowAction);
els.editorColumnSearch?.addEventListener("input", handlers.handleEditorColumnSearchChanged);
els.editorColumnSearchResults?.addEventListener("click", handlers.handleEditorColumnSearchResultClicked);

els.editorFixSearch?.addEventListener("input", handlers.handleEditorFixSearchChanged);
els.editorFixResults?.addEventListener("click", handlers.handleEditorFixResultClicked);
els.editorGradeSave?.addEventListener("click", handlers.handleEditorGradeSaveClicked);
els.editorDelimiterFilter?.addEventListener("change", handlers.handleDelimiterFilterChanged);

// Wizard navigation
els.btnWizardPrev?.addEventListener("click", handlers.handleWizardPrev);
els.btnWizardNext?.addEventListener("click", handlers.handleWizardNext);
els.btnWizardFinish?.addEventListener("click", handlers.handleWizardFinish);

// OCR wiring
els.ocrImageUpload?.addEventListener("change", handlers.handleOcrImageUpload);
els.btnOcrProcess?.addEventListener("click", handlers.handleOcrProcess);
els.btnOcrApprove?.addEventListener("click", handlers.handleOcrApprove);
els.btnOcrReset?.addEventListener("click", handlers.handleOcrReset);

// Initial state
switchView('inputs');
try {
  handlers.updateWizardUI?.();
  // Initialize input method UI (show file input by default)
  if (els.editorInputFileContainer) {
    els.editorInputFileContainer.style.display = "block";
  }
  if (els.editorInputTextareaContainer) {
    els.editorInputTextareaContainer.style.display = "none";
  }
} catch (e) {
  console.error("Error initializing wizard UI:", e);
}

// Footer year
const footerYearEl = document.getElementById("footerYear");
if (footerYearEl) footerYearEl.textContent = String(new Date().getFullYear());

// Splash screen fade-out
const splashScreen = document.getElementById("splashScreen");
if (splashScreen) {
  // Fade out after 1.5 seconds
  setTimeout(() => {
    splashScreen.classList.add("fade-out");
    // Remove from DOM after animation completes
    setTimeout(() => {
      splashScreen.remove();
    }, 350); // Match transition-slow duration
  }, 1500);
}

