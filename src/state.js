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
      inputFileName: "",
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


