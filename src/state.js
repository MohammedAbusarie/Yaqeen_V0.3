/**
 * App state container + reset logic (internal).
 */

/**
 * @typedef {Object} AppState
 * @property {ArrayBuffer|null} workbookArrayBuffer
 * @property {string|null} workbookName
 * @property {any} editor
 */

/**
 * @returns {AppState}
 */
export function createInitialState() {
  return {
    workbookArrayBuffer: null,
    workbookName: null,
    editor: {
      // file
      workbookLoaded: false,
      workbookSheetNames: [],

      // selection
      scopeMode: "single", // 'single' | 'multi'
      selectedSheetName: "",
      selectedColumnKey: "",
      taskType: "attendance", // 'attendance' | 'grade'

      // input + preview
      inputMethod: "file", // 'file' | 'textarea'
      inputFileName: "",
      inputTextContent: null, // For textarea input
      originalInputData: null, // Store original parsed data for download
      previewRows: null,
      columnMap: null,
      selectedColumn: null,
      orderedEntries: null, // For preserving delimiter/title information in ordered view
      columnOptions: null, // For column search functionality
      // user edits in preview
      manualEditsByIndex: {}, // preview index -> partial overrides
      
      // highlight settings for download
      highlightEnabled: true, // default: highlight enabled
      highlightColor: "#FFFF00", // default: yellow
      
      // wizard
      wizardStep: 1, // 1-4
    },
    ocr: {
      // OCR experimental feature state
      uploadedImages: [],
      processingResults: null,
      approvedIds: [],
      currentStep: 1, // 1: upload, 2: processing, 3: review
    },
    sheetMerger: {
      // Sheet Merger feature state
      workbookArrayBuffer: null,
      workbookName: null,
      workbookLoaded: false,
      workbookSheetNames: [],
      allColumns: [], // {sheet, columnIndex, headerText, headerRow, sampleValues}
      mapping: {}, // {[sheetName]: {[position]: columnKey}}
      mergedData: null,
      previewRows: [],
      previewRowsLoaded: 100, // Number of rows to display in preview (starts at 100)
      eliminateHeaders: false,
      expandedSheets: [], // Track which sheet groups are expanded
      searchQuery: "", // Current search filter text
      maxPositions: 10, // Dynamic column count, start with 10
      sheetColors: {}, // Map of sheet name to color for visual identification
    },
  };
}

/**
 * Resets all inputs and state back to initial.
 * @param {any} els
 * @param {(msg: string, kind?: 'info'|'ok'|'error') => void} setStatusFn
 * @param {AppState} state
 */
export function resetAll(els, setStatusFn, state) {
  state.workbookArrayBuffer = null;
  state.workbookName = null;
  const initialEditor = createInitialState().editor;
  state.editor = { ...initialEditor, wizardStep: 1 };

  if (els.btnDownloadJson) els.btnDownloadJson.disabled = true;
  if (els.btnDownloadTxt) els.btnDownloadTxt.disabled = true;
  if (els.btnDownloadPdf) els.btnDownloadPdf.disabled = true;
  if (els.btnEditorDownload) els.btnEditorDownload.disabled = true;
  if (setStatusFn) setStatusFn("");
}


