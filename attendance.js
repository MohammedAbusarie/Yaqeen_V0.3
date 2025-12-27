// Core attendance processing logic (ported from gui/core/attendance.py)
// Runs fully in the browser. Requires global XLSX from SheetJS.

/**
 * @typedef {Object} FoundLogEntry
 * @property {string} sheet
 * @property {string} id
 * @property {string} name
 * @property {string} target_cell
 * @property {string} col_letter
 * @property {number} header_row
 * @property {number} first_data_row
 * @property {number} last_data_row
 */

/**
 * @typedef {Object} ParsedStudentIds
 * @property {Array<{type:'id', id:string, section?:number} | {type:'title', title:string}>} orderedEntries
 * @property {Set<string>} targetIdsSet
 * @property {number} totalLoadedUnique
 * @property {Record<string, number>} idCounts
 * @property {Record<number, Record<string, number>>} sectionIdCounts
 */

export class ValidationError extends Error {}
export class DownloadError extends Error {}
export class FileError extends Error {}
export class ProcessingError extends Error {}

// -----------------------------
// Editor types (generic column edits)
// -----------------------------
/**
 * @typedef {Object} ColumnOption
 * @property {string} key         // stable unique key: `${kind}::${headerText}`
 * @property {string} headerText  // as shown in the workbook header cell
 * @property {'lecture'|'section'|'unknown'} kind
 * @property {number} occurrences
 * @property {Array<{ sheet: string, header_row: number, col1: number, col_letter: string }>} locations
 */

/**
 * @typedef {Object} EditorPreviewRow
 * @property {number} index          // 1-based index in the input file order
 * @property {string} input_id
 * @property {string} sheet
 * @property {number|null} row_index1
 * @property {string} student_id
 * @property {string} student_name
 * @property {string} cell
 * @property {string} col_letter
 * @property {string|number} old_value
 * @property {string|number} new_value
 * @property {'matched'|'notFound'|'ambiguous'|'manuallyFixed'} match_status
 * @property {string} [note]
 */

/**
 * @typedef {Object} ColumnMapEntry
 * @property {string} sheet
 * @property {string} header_text
 * @property {'lecture'|'section'|'unknown'} kind
 * @property {number} header_row
 * @property {string} col_letter
 */

// -----------------------------
// ID normalization + parsing
// -----------------------------
export function normalizeId(idValue) {
  if (idValue === null || idValue === undefined) return null;
  if (typeof idValue === "number" && Number.isFinite(idValue)) return String(Math.trunc(idValue));
  const cleaned = String(idValue).trim().replace(/\.0$/, "");
  return cleaned;
}

export function parseStudentIdsText(text) {
  if (typeof text !== "string") throw new FileError("Student IDs file content is invalid");

  // Mirrors gui/core/attendance.py load_student_ids()
  // ordered_entries contains:
  // - { type: "id", id: "...", section?: number }
  // - { type: "title", title: "..." }
  const orderedEntries = [];
  const orderedIdsOnly = [];
  const idCounts = {}; // global count in whole file (unique semantics match GUI json)
  const sectionIdCounts = {}; // section -> { id -> count }

  let currentSectionId = null;
  let hasSeenFirstTitle = false;

  const lines = text.replace(/\r\n/g, "\n").split("\n");
  for (const rawLine of lines) {
    const cleaned = rawLine.trim();
    if (!cleaned) {
      if (hasSeenFirstTitle) currentSectionId = null;
      continue;
    }

    const isValidId = /^\d+$/.test(cleaned) && cleaned.length > 5;
    if (isValidId) {
      const sid = normalizeId(cleaned);
      const entry = { type: "id", id: sid };
      if (currentSectionId !== null) {
        entry.section = currentSectionId;
        if (!sectionIdCounts[currentSectionId]) sectionIdCounts[currentSectionId] = {};
        sectionIdCounts[currentSectionId][sid] = (sectionIdCounts[currentSectionId][sid] || 0) + 1;
      }
      orderedEntries.push(entry);
      orderedIdsOnly.push(sid);
      idCounts[sid] = (idCounts[sid] || 0) + 1;
    } else {
      orderedEntries.push({ type: "title", title: cleaned });
      hasSeenFirstTitle = true;
      currentSectionId = Object.keys(sectionIdCounts).length + 1;
      if (!sectionIdCounts[currentSectionId]) sectionIdCounts[currentSectionId] = {};
    }
  }

  const idsSet = new Set(orderedIdsOnly);
  if (idsSet.size === 0) {
    throw new FileError("No valid student IDs found. IDs must be numeric and at least 6 digits.");
  }

  return {
    orderedEntries,
    targetIdsSet: idsSet,
    totalLoadedUnique: idsSet.size,
    idCounts,
    sectionIdCounts,
  };
}

// -----------------------------
// XLSX loading + Google Sheet URL download
// -----------------------------
function assertXlsxLoaded() {
  // SheetJS exposes global XLSX
  if (typeof window === "undefined" || !window.XLSX) {
    throw new ProcessingError("XLSX parser not loaded. Please refresh the page.");
  }
}

export function googleSheetEditUrlToExportXlsxUrl(url) {
  if (!url || !String(url).includes("docs.google.com/spreadsheets")) {
    throw new ValidationError(
      "Invalid Google Sheet URL. Must contain 'docs.google.com/spreadsheets'."
    );
  }
  const u = String(url);
  if (u.includes("/edit")) {
    const base = u.split("?")[0].replace("/edit", "/export");
    return `${base}?format=xlsx`;
  }
  return u;
}

export function sharePointUrlToDownloadUrl(url) {
  if (!url || !String(url).includes("sharepoint.com")) {
    throw new ValidationError(
      "Invalid SharePoint URL. Must contain 'sharepoint.com'."
    );
  }
  
  const u = String(url);
  
  // Extract file ID from SharePoint URL
  // Format: https://[tenant]-my.sharepoint.com/:x:/g/personal/[user]/[fileid]?e=[token] or ?rtime=...
  const fileIdMatch = u.match(/\/personal\/[^\/]+\/([^?\/]+)/);
  if (!fileIdMatch) {
    throw new ValidationError(
      "Could not extract file ID from SharePoint URL. Please ensure the URL is in the correct format."
    );
  }
  
  const fileId = fileIdMatch[1];
  
  // Extract tenant and user path
  const tenantMatch = u.match(/https:\/\/([^-]+)-my\.sharepoint\.com/);
  if (!tenantMatch) {
    throw new ValidationError(
      "Could not extract tenant from SharePoint URL."
    );
  }
  
  const tenant = tenantMatch[1];
  const userPathMatch = u.match(/\/personal\/([^\/]+)/);
  const userPath = userPathMatch ? userPathMatch[1] : "";
  
  // For SharePoint/OneDrive files shared with "Anyone with the link",
  // we need to use a format that works without authentication.
  // The best approach is to convert the sharing link to a direct download format.
  
  // Try format: Convert sharing link to download format using webUrl parameter
  // This works for files shared publicly without requiring authentication
  // Format: https://[tenant]-my.sharepoint.com/personal/[user]/_layouts/15/download.aspx?UniqueId=[fileid]&download=1
  const downloadUrl = `https://${tenant}-my.sharepoint.com/personal/${userPath}/_layouts/15/download.aspx?UniqueId=${fileId}&download=1`;
  
  return downloadUrl;
}

export async function fetchXlsxFromUrl(url) {
  const urlStr = String(url).trim();
  let exportUrl;
  let urlType;
  
  // Detect URL type
  if (urlStr.includes("docs.google.com/spreadsheets")) {
    urlType = "google";
    exportUrl = googleSheetEditUrlToExportXlsxUrl(urlStr);
  } else if (urlStr.includes("sharepoint.com")) {
    urlType = "sharepoint";
    exportUrl = sharePointUrlToDownloadUrl(urlStr);
  } else {
    throw new ValidationError(
      "Unsupported URL type. Please provide a Google Sheets or SharePoint URL, or upload the file directly."
    );
  }
  
  let res;
  try {
    res = await fetch(exportUrl, { method: "GET", mode: "cors" });
  } catch (e) {
    const errorMsg = urlType === "google"
      ? "Could not download the sheet (network/CORS). If this is a Google Sheets CORS restriction, export as .xlsx manually and upload instead."
      : "Could not download the sheet (network/CORS). SharePoint files often require authentication. Please download the file manually and upload it instead.";
    throw new DownloadError(errorMsg);
  }
  
  if (!res.ok) {
    if (urlType === "sharepoint" && res.status === 403) {
      throw new DownloadError(
        "SharePoint returned 403 Forbidden. Even though the file may be accessible in your browser when you're logged in, SharePoint's API endpoints require authentication tokens that cannot be provided from this web app. This is a limitation of SharePoint's security model - programmatic access requires OAuth authentication. Please download the file manually from SharePoint (right-click the file → Download) and upload it using the file upload option above."
      );
    }
    const errorMsg = urlType === "google"
      ? `Could not download the sheet (HTTP ${res.status}). Make sure the sheet is shared with 'Anyone with the link' and try exporting to .xlsx manually if needed.`
      : `Could not download the sheet (HTTP ${res.status}). SharePoint files require authentication for programmatic access. Please download the file manually from SharePoint and upload it instead.`;
    throw new DownloadError(errorMsg);
  }
  
  // Validate that we got a binary Excel file, not HTML
  const contentType = res.headers.get("content-type") || "";
  const isExcelFile = contentType.includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") ||
                      contentType.includes("application/octet-stream") ||
                      contentType.includes("application/vnd.ms-excel");
  
  // Check if response might be HTML (common SharePoint error pages)
  if (!isExcelFile && (contentType.includes("text/html") || contentType.includes("application/json"))) {
    throw new DownloadError(
      "SharePoint returned an HTML page instead of the Excel file. This usually means the file requires authentication. Please download the file manually from SharePoint and upload it instead."
    );
  }
  
  const arrayBuffer = await res.arrayBuffer();
  
  // Additional validation: Check if the file starts with Excel signature (ZIP file signature)
  // Excel files are ZIP archives, so they start with PK (0x504B)
  const view = new Uint8Array(arrayBuffer);
  if (view.length < 2 || view[0] !== 0x50 || view[1] !== 0x4B) {
    // Might be HTML - check for common HTML tags
    const textDecoder = new TextDecoder();
    const textStart = textDecoder.decode(view.slice(0, Math.min(500, view.length)));
    if (textStart.includes("<!DOCTYPE") || textStart.includes("<html") || textStart.includes("<HTML")) {
      throw new DownloadError(
        "SharePoint returned an HTML page (likely a login or error page) instead of the Excel file. SharePoint files typically require authentication. Please download the file manually from SharePoint and upload it instead."
      );
    }
  }
  
  return arrayBuffer;
}

// Keep the old function for backward compatibility
export async function fetchXlsxFromGoogleSheetUrl(url) {
  return fetchXlsxFromUrl(url);
}

export function readWorkbookFromArrayBuffer(arrayBuffer) {
  assertXlsxLoaded();
  
  // Try reading with standard options first
  try {
    const wb = window.XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: false,
      defval: "", // Default value for empty cells
      cellFormula: false, // Ignore formulas, read calculated values only
      cellStyles: false, // Ignore cell styles to avoid parsing issues
    });
    return wb;
  } catch (error) {
    const errorMsg = error?.message || String(error);
    
    // If we get an "Unsupported value type" error, try with raw mode
    if (errorMsg.includes("Unsupported value type") || errorMsg.includes("unsupported")) {
      try {
        // Try reading with raw mode enabled - this can help with ODS files
        const wb = window.XLSX.read(arrayBuffer, {
          type: "array",
          cellDates: false,
          defval: "",
          raw: true, // Use raw values to avoid type conversion issues
          cellFormula: false, // Ignore formulas, read calculated values
          cellStyles: false, // Ignore styles
        });
        
        // Convert raw values to strings where needed (only process existing cells)
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          if (!ws) continue;
          
          // Only process cells that actually exist (non-empty cells)
          for (const cellAddress in ws) {
            // Skip special properties that start with !
            if (cellAddress.startsWith("!")) continue;
            
            const cell = ws[cellAddress];
            if (!cell) continue;
            
            // If cell has a formula but no calculated value, try to extract from w attribute
            if (cell.f && (cell.v === undefined || cell.v === null)) {
              // Try to get the calculated value from the w (formatted text) attribute
              if (cell.w !== undefined && cell.w !== null) {
                // Use the formatted text as fallback
                cell.v = cell.w;
                cell.t = "s";
              } else {
                // If no calculated value available, set to empty
                cell.v = "";
                cell.t = "s";
              }
            }
            
            // Ensure cell has a valid value
            if (cell.v === undefined || cell.v === null) {
              cell.v = "";
              cell.t = "s";
            } else if (typeof cell.v !== "string" && typeof cell.v !== "number" && typeof cell.v !== "boolean") {
              // Convert unexpected types to string
              cell.v = String(cell.v);
              cell.t = "s";
            }
          }
        }
        
        return wb;
      } catch (retryError) {
        // If retry also fails, provide helpful error message
        throw new FileError(
          "The ODS file contains data types that cannot be processed. " +
          "This often happens with formulas or special cell formats. " +
          "Please try: (1) Opening the file in LibreOffice Calc and saving as .xlsx format, " +
          "or (2) Exporting to CSV format, or (3) Removing any formulas and replacing them with their calculated values."
        );
      }
    }
    
    // For other errors, provide generic error message
    throw new FileError(`Failed to read the spreadsheet file: ${errorMsg}`);
  }
}

// -----------------------------
// Worksheet scan helpers
// -----------------------------
function getSheetRange(ws) {
  const ref = ws["!ref"];
  if (!ref) return null;
  return window.XLSX.utils.decode_range(ref);
}

function cellValue(ws, r1, c1) {
  // r1,c1 are 1-based
  const addr = window.XLSX.utils.encode_cell({ r: r1 - 1, c: c1 - 1 });
  const cell = ws[addr];
  if (!cell) return null;
  
  // Handle edge cases from ODS files or unusual cell types
  const v = cell.v;
  if (v === undefined || v === null) return null;
  
  // Ensure we return a valid primitive type
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") {
    return v;
  }
  
  // For unexpected types, convert to string
  return String(v);
}

function findColByTextInRow(ws, rowIdx1, textLower) {
  const range = getSheetRange(ws);
  if (!range) return null;

  for (let c0 = range.s.c; c0 <= range.e.c; c0++) {
    const v = cellValue(ws, rowIdx1, c0 + 1);
    if (v === null || v === undefined) continue;
    const s = String(v).toLowerCase();
    if (s.includes(textLower)) return c0 + 1; // 1-based
  }
  return null;
}

/**
 * Detects the ID column and name column(s) in a worksheet.
 * Name columns are detected by scanning the first 5-10 columns for strings (letters).
 * ID column is detected by matching values against target IDs (handling email format).
 * 
 * @param {object} ws - Worksheet object from SheetJS
 * @param {Set<string>|null} targetIdsSet - Set of target student IDs to match against. If null, ID detection is skipped.
 * @param {number} startRow1 - Starting row to scan (default: 2)
 * @param {number} maxRowsToScan - Maximum number of rows to scan (default: 50)
 * @returns {{idCol: number|null, nameCol: number, nameCol2: number|null}} - 1-based column indices
 */
function detectIdAndNameColumns(ws, targetIdsSet = null, startRow1 = 2, maxRowsToScan = 50) {
  const range = getSheetRange(ws);
  if (!range) return { idCol: null, nameCol: 1, nameCol2: null };
  
  const maxRow1 = Math.min(range.e.r + 1, startRow1 + maxRowsToScan - 1);
  const maxCol1 = range.e.c + 1;
  
  // Helper to extract ID from value (handles email format)
  function extractId(value, targetSet) {
    if (!targetSet || targetSet.size === 0) return null;
    
    const normalized = normalizeId(value);
    if (!normalized) return null;
    
    // Check if it contains "@" (email format)
    if (normalized.includes("@")) {
      const username = normalized.split("@")[0];
      if (targetSet.has(username)) {
        return username;
      }
    }
    
    // Direct match
    if (targetSet.has(normalized)) return normalized;
    
    return null;
  }
  
  // Helper to check if a value looks like a name (string with letters, not ID-like)
  function looksLikeName(value) {
    if (value === null || value === undefined) return false;
    const str = String(value).trim();
    if (str.length === 0) return false;
    
    // Must contain letters (not just numbers)
    if (!/[a-zA-Z]/.test(str)) return false;
    
    // Should not look like an ID (numeric or email with numeric prefix)
    if (/^\d+$/.test(str) && str.length > 5) return false;
    if (/^\d{6,}@/.test(str)) return false;
    
    return true;
  }
  
  // 1. Detect name columns (scan first 5-10 columns)
  const nameColumnScanLimit = Math.min(10, maxCol1);
  /** @type {Map<number, number>} */ // column -> text count
  const nameColumnScores = new Map();
  
  for (let c = 1; c <= nameColumnScanLimit; c++) {
    let textCount = 0;
    let totalRows = 0;
    
    for (let r = startRow1; r <= maxRow1; r++) {
      const value = cellValue(ws, r, c);
      totalRows++;
      if (looksLikeName(value)) {
        textCount++;
      }
    }
    
    // Consider it a name column if >50% of rows have text-like values
    if (totalRows > 0 && textCount / totalRows > 0.5) {
      nameColumnScores.set(c, textCount);
    }
  }
  
  // Find name columns (prefer adjacent pairs)
  let nameCol = 1;
  let nameCol2 = null;
  
  if (nameColumnScores.size > 0) {
    const nameCols = Array.from(nameColumnScores.keys()).sort((a, b) => a - b);
    
    // Look for adjacent pairs
    for (let i = 0; i < nameCols.length - 1; i++) {
      if (nameCols[i + 1] === nameCols[i] + 1) {
        nameCol = nameCols[i];
        nameCol2 = nameCols[i + 1];
        break;
      }
    }
    
    // If no adjacent pair found, use the first name column
    if (nameCol2 === null && nameCols.length > 0) {
      nameCol = nameCols[0];
    }
  }
  
  // 2. Detect ID column (scan all columns)
  let idCol = null;
  /** @type {Map<number, number>} */ // column -> match count
  const idColumnScores = new Map();
  
  // Helper to check if a value looks like an ID (numeric or email with numeric username)
  function looksLikeId(value) {
    if (value === null || value === undefined) return false;
    const normalized = normalizeId(value);
    if (!normalized) return false;
    
    // Check for email format with numeric username
    if (normalized.includes("@")) {
      const username = normalized.split("@")[0];
      if (/^\d+$/.test(username) && username.length > 5) {
        return true;
      }
    }
    
    // Check for pure numeric ID
    if (/^\d+$/.test(normalized) && normalized.length > 5) {
      return true;
    }
    
    return false;
  }
  
  if (targetIdsSet && targetIdsSet.size > 0) {
    // With target IDs: scan all columns and count matches
    for (let c = 1; c <= maxCol1; c++) {
      // Skip columns that are already identified as name columns
      if (c === nameCol || c === nameCol2) continue;
      
      let matchCount = 0;
      let totalRows = 0;
      
      for (let r = startRow1; r <= maxRow1; r++) {
        const value = cellValue(ws, r, c);
        totalRows++;
        if (extractId(value, targetIdsSet)) {
          matchCount++;
        }
      }
      
      // Consider it an ID column if it has at least one match
      // Prefer columns with higher match rates
      if (matchCount > 0) {
        idColumnScores.set(c, matchCount);
      }
    }
    
    // Select the column with the most matches
    if (idColumnScores.size > 0) {
      let bestCol = null;
      let bestScore = 0;
      for (const [col, score] of idColumnScores) {
        if (score > bestScore) {
          bestScore = score;
          bestCol = col;
        }
      }
      idCol = bestCol;
    }
  } else {
    // Without target IDs: use pattern-based detection
    // Look for columns with email patterns or numeric IDs
    for (let c = 1; c <= maxCol1; c++) {
      // Skip columns that are already identified as name columns
      if (c === nameCol || c === nameCol2) continue;
      
      let idLikeCount = 0;
      let totalRows = 0;
      
      for (let r = startRow1; r <= maxRow1; r++) {
        const value = cellValue(ws, r, c);
        totalRows++;
        if (looksLikeId(value)) {
          idLikeCount++;
        }
      }
      
      // Consider it an ID column if >50% of rows look like IDs
      if (totalRows > 0 && idLikeCount / totalRows > 0.5) {
        idColumnScores.set(c, idLikeCount);
      }
    }
    
    // Select the column with the most ID-like values
    if (idColumnScores.size > 0) {
      let bestCol = null;
      let bestScore = 0;
      for (const [col, score] of idColumnScores) {
        if (score > bestScore) {
          bestScore = score;
          bestCol = col;
        }
      }
      idCol = bestCol;
    }
  }
  
  return { idCol, nameCol, nameCol2 };
}

export function detectLectureSectionBounds(ws) {
  // Returns the boundary columns used in this workbook template, if present.
  // col_* are 1-based.
  const colSection = findColByTextInRow(ws, 1, "attendance section");
  const colLecture = findColByTextInRow(ws, 1, "attendance lecture");
  return { colSection, colLecture };
}

function getSearchBoundsForKind({ kind, colSection, colLecture, maxCol1 }) {
  if (!colSection) return { startCol1: 1, endCol1: maxCol1, kind: "unknown" };
  if (kind === "section") {
    return { startCol1: colSection, endCol1: colLecture ? colLecture - 1 : maxCol1, kind: "section" };
  }
  if (kind === "lecture") {
    if (!colLecture) return { startCol1: 1, endCol1: 0, kind: "lecture" }; // empty range
    // Based on actual Excel structure analysis: lecture week columns (W1, W2, W3...) start
    // at the same column where "Attendance Lecture" header is found in row 1.
    // Section week columns end at colLecture - 1, so lecture starts at colLecture.
    return { startCol1: colLecture, endCol1: maxCol1, kind: "lecture" };
  }
  // unknown: scan everything, but keep kind as unknown
  return { startCol1: 1, endCol1: maxCol1, kind: "unknown" };
}

export function listColumnOptions(workbook, scope) {
  assertXlsxLoaded();
  if (!workbook || !Array.isArray(workbook.SheetNames)) throw new ProcessingError("Invalid workbook.");

  const scopeMode = scope?.mode === "single" ? "single" : "multi";
  const selectedSheet = String(scope?.sheetName || "");
  const sheetNames =
    scopeMode === "single"
      ? [selectedSheet].filter(Boolean)
      : workbook.SheetNames.slice();

  if (sheetNames.length === 0) throw new ValidationError("No sheets selected.");

  /** @type {Map<string, ColumnOption>} */
  const byKey = new Map();

  for (const sheetName of sheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;
    const range = getSheetRange(ws);
    if (!range) continue;
    const maxCol1 = range.e.c + 1;

    const { colSection, colLecture } = detectLectureSectionBounds(ws);

    /** @type {Array<'section'|'lecture'|'unknown'>} */
    const kindsToScan = colSection ? ["section", "lecture"] : ["unknown"];

    // First, collect all headers by their text (case-insensitive) to merge row 1 with rows 2-5
    /** @type {Map<string, { headerText: string, kind: string, locations: Array }>} */
    const byHeaderText = new Map();

    // Scan row 1 first (always "unknown" kind)
    for (let c = 1; c <= maxCol1; c++) {
      const raw = cellValue(ws, 1, c);
      const headerText = String(raw ?? "").trim();
      if (!headerText) continue;
      const headerLower = headerText.toLowerCase();
      const colLetter = window.XLSX.utils.encode_col(c - 1);
      const loc = { sheet: sheetName, header_row: 1, col1: c, col_letter: colLetter };
      
      if (!byHeaderText.has(headerLower)) {
        byHeaderText.set(headerLower, {
          headerText,
          kind: "unknown",
          locations: [loc],
        });
      } else {
        byHeaderText.get(headerLower).locations.push(loc);
      }
    }

    // Then scan rows 2-5 with detected kinds
    for (const kind of kindsToScan) {
      const { startCol1, endCol1 } = getSearchBoundsForKind({ kind, colSection, colLecture, maxCol1 });
      if (endCol1 < startCol1) continue;

      for (let r = 2; r <= 5; r++) {
        // stop early if we reached data rows (ID-like in col B)
        const idCheck = normalizeId(cellValue(ws, r, 2));
        if (idCheck && /^\d+$/.test(idCheck) && idCheck.length > 5) break;

        for (let c = startCol1; c <= endCol1; c++) {
          const raw = cellValue(ws, r, c);
          const headerText = String(raw ?? "").trim();
          if (!headerText) continue;
          const headerLower = headerText.toLowerCase();
          const colLetter = window.XLSX.utils.encode_col(c - 1);
          const loc = { sheet: sheetName, header_row: r, col1: c, col_letter: colLetter };

          if (!byHeaderText.has(headerLower)) {
            byHeaderText.set(headerLower, {
              headerText,
              kind,
              locations: [loc],
            });
          } else {
            const entry = byHeaderText.get(headerLower);
            entry.locations.push(loc);
            // Upgrade from unknown to detected kind if found in rows 2-5
            // Note: We'll split by column ranges in the final conversion step,
            // so we don't need to worry about section vs lecture conflicts here
            if (entry.kind === "unknown" && kind !== "unknown") {
              entry.kind = kind;
            }
          }
        }
      }
    }

    // Now convert to byKey format, splitting locations by their actual column ranges
    // This ensures that headers appearing in both section and lecture areas get separate entries
    for (const [headerLower, entry] of byHeaderText) {
      if (!colSection || !colLecture) {
        // No section/lecture boundaries detected, use entry as-is
        const key = `${entry.kind}::${entry.headerText}`;
        if (!byKey.has(key)) {
          byKey.set(key, {
            key,
            headerText: entry.headerText,
            kind: entry.kind,
            occurrences: entry.locations.length,
            locations: entry.locations,
          });
        } else {
          const opt = byKey.get(key);
          opt.occurrences += entry.locations.length;
          opt.locations.push(...entry.locations);
        }
        continue;
      }

      // Split locations by section vs lecture column ranges
      const sectionLocations = [];
      const lectureLocations = [];
      const unknownLocations = [];

      for (const loc of entry.locations) {
        // Determine which range this location falls into
        // Section: colSection to colLecture - 1
        // Lecture: colLecture to maxCol1
        if (loc.col1 >= colSection && loc.col1 < colLecture) {
          sectionLocations.push(loc);
        } else if (loc.col1 >= colLecture) {
          lectureLocations.push(loc);
        } else {
          unknownLocations.push(loc);
        }
      }

      // Create separate entries for section and lecture if they have locations
      if (sectionLocations.length > 0) {
        const key = `section::${entry.headerText}`;
        if (!byKey.has(key)) {
          byKey.set(key, {
            key,
            headerText: entry.headerText,
            kind: "section",
            occurrences: sectionLocations.length,
            locations: sectionLocations,
          });
        } else {
          const opt = byKey.get(key);
          opt.occurrences += sectionLocations.length;
          opt.locations.push(...sectionLocations);
        }
      }

      if (lectureLocations.length > 0) {
        const key = `lecture::${entry.headerText}`;
        if (!byKey.has(key)) {
          byKey.set(key, {
            key,
            headerText: entry.headerText,
            kind: "lecture",
            occurrences: lectureLocations.length,
            locations: lectureLocations,
          });
        } else {
          const opt = byKey.get(key);
          opt.occurrences += lectureLocations.length;
          opt.locations.push(...lectureLocations);
        }
      }

      // Handle unknown locations (outside both ranges)
      if (unknownLocations.length > 0) {
        const key = `unknown::${entry.headerText}`;
        if (!byKey.has(key)) {
          byKey.set(key, {
            key,
            headerText: entry.headerText,
            kind: "unknown",
            occurrences: unknownLocations.length,
            locations: unknownLocations,
          });
        } else {
          const opt = byKey.get(key);
          opt.occurrences += unknownLocations.length;
          opt.locations.push(...unknownLocations);
        }
      }
    }
  }

  return Array.from(byKey.values()).sort((a, b) => {
    const k1 = `${a.kind}::${a.headerText}`.toLowerCase();
    const k2 = `${b.kind}::${b.headerText}`.toLowerCase();
    return k1.localeCompare(k2);
  });
}

function buildStudentIndexForSheet(ws, targetIdsSet = null) {
  const range = getSheetRange(ws);
  if (!range) return { byId: new Map(), rows: [] };
  const maxRow1 = range.e.r + 1;

  /** @type {Map<string, Array<{row1:number, id:string, name:string}>>} */
  const byId = new Map();
  /** @type {Array<{row1:number, id:string, name:string}>>} */
  const rows = [];

  // Detect ID and name columns dynamically (once, before the loop)
  const { idCol, nameCol, nameCol2 } = detectIdAndNameColumns(ws, targetIdsSet);
  
  // Smart fallback: if ID column not detected, try to find a column with email/numeric patterns
  let finalIdCol = idCol;
  if (!finalIdCol) {
    const range = getSheetRange(ws);
    if (range) {
      const maxCol1 = range.e.c + 1;
      const maxRow1 = Math.min(range.e.r + 1, 52); // Scan first 50 data rows
      
      // Look for a column that contains emails or numeric IDs (excluding name columns)
      for (let c = 1; c <= maxCol1; c++) {
        if (c === nameCol || c === nameCol2) continue; // Skip name columns
        
        let idLikeCount = 0;
        let totalRows = 0;
        
        for (let r = 2; r <= maxRow1; r++) {
          const value = cellValue(ws, r, c);
          totalRows++;
          if (value !== null && value !== undefined) {
            const normalized = normalizeId(value);
            if (normalized) {
              // Check for email with numeric username or pure numeric ID
              if (normalized.includes("@")) {
                const username = normalized.split("@")[0];
                if (/^\d+$/.test(username) && username.length > 5) {
                  idLikeCount++;
                }
              } else if (/^\d+$/.test(normalized) && normalized.length > 5) {
                idLikeCount++;
              }
            }
          }
        }
        
        // If >50% of rows look like IDs, use this column
        if (totalRows > 0 && idLikeCount / totalRows > 0.5) {
          finalIdCol = c;
          break;
        }
      }
    }
    
    // Last resort: fallback to column 2 (original assumption)
    if (!finalIdCol) {
      finalIdCol = 2;
    }
  }

  for (let r = 2; r <= maxRow1; r++) {
    const rawId = cellValue(ws, r, finalIdCol);
    
    // Extract ID, handling email format
    let sid = null;
    if (rawId !== null && rawId !== undefined) {
      const normalized = normalizeId(rawId);
      
      if (normalized) {
        // Check if it contains "@" (email format)
        if (normalized.includes("@")) {
          const username = normalized.split("@")[0];
          if (targetIdsSet && targetIdsSet.has(username)) {
            sid = username;
          } else if (!targetIdsSet && /^\d+$/.test(username) && username.length > 5) {
            // Pattern-based: if username is numeric and long enough, use it
            sid = username;
          }
        } else {
          // Direct match
          if (targetIdsSet && targetIdsSet.has(normalized)) {
            sid = normalized;
          } else if (!targetIdsSet && /^\d+$/.test(normalized) && normalized.length > 5) {
            // Pattern-based: pure numeric ID
            sid = normalized;
          }
        }
      }
    }
    
    if (!sid) continue;
    
    // Build name from detected name column(s)
    let name = "";
    if (nameCol2 !== null) {
      // Concatenate two name columns
      const name1 = String(cellValue(ws, r, nameCol) ?? "").trim();
      const name2 = String(cellValue(ws, r, nameCol2) ?? "").trim();
      name = `${name1} ${name2}`.trim();
    } else {
      // Single name column
      name = String(cellValue(ws, r, nameCol) ?? "").trim();
    }
    
    const row = { row1: r, id: sid, name };
    rows.push(row);
    if (!byId.has(sid)) byId.set(sid, [row]);
    else byId.get(sid).push(row);
  }

  return { byId, rows };
}

export function buildStudentSearchRows(workbook, scope) {
  assertXlsxLoaded();
  if (!workbook || !Array.isArray(workbook.SheetNames)) throw new ProcessingError("Invalid workbook.");
  const scopeMode = scope?.mode === "single" ? "single" : "multi";
  const selectedSheet = String(scope?.sheetName || "");
  const sheetNames =
    scopeMode === "single"
      ? [selectedSheet].filter(Boolean)
      : workbook.SheetNames.slice();

  /** @type {Array<{sheet:string,row1:number,id:string,name:string}>} */
  const out = [];
  for (const sheetName of sheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;
    const index = buildStudentIndexForSheet(ws);
    for (const r of index.rows) out.push({ sheet: sheetName, row1: r.row1, id: r.id, name: r.name });
  }
  return out;
}

export function parseGradesText(text) {
  if (typeof text !== "string") throw new FileError("Grades file content is invalid");
  const lines = text.replace(/\r\n/g, "\n").split("\n");
  /** @type {Array<{ id: string, grade: string }>} */
  const rows = [];
  /** @type {Array<{ type: "id", id: string, grade: string } | { type: "title", title: string }>} */
  const orderedEntries = [];

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) {
      // Empty line resets delimiter context (same as parseStudentIdsText)
      continue;
    }
    
    const parts = line.split(",").map((p) => p.trim());
    if (parts.length < 2) {
      // No comma - this is a delimiter/title line
      orderedEntries.push({ type: "title", title: line });
      continue;
    }
    
    const sid = normalizeId(parts[0]);
    // Check if first part is a valid ID (numeric, 6+ digits)
    if (!sid || !/^\d+$/.test(sid) || sid.length <= 5) {
      // Not a valid ID, treat this line as a delimiter/title
      orderedEntries.push({ type: "title", title: line });
      continue;
    }
    
    const grade = parts.slice(1).join(",").trim(); // allow commas in grade text
    if (!grade) {
      // Valid ID but no grade - this is an error
      throw new FileError(`Missing grade for ID '${sid}'`);
    }
    
    const entry = { type: "id", id: sid, grade };
    rows.push({ id: sid, grade });
    orderedEntries.push(entry);
  }

  if (rows.length === 0) throw new FileError("No valid grade rows found.");
  
  return {
    rows,
    orderedEntries,
  };
}

export function computeEditorPreview({
  workbook,
  scope,
  columnKey,
  taskType,
  orderedAttendanceIds,
  attendanceIdsSet,
  gradesRows,
}) {
  assertXlsxLoaded();
  if (!workbook || !Array.isArray(workbook.SheetNames)) throw new ProcessingError("Invalid workbook.");

  const opts = listColumnOptions(workbook, scope);
  let selected = opts.find((o) => o.key === columnKey) || null;
  
  // If exact key match not found, try to find by header text (handles row 1 vs rows 2-5 key mismatch)
  if (!selected && columnKey) {
    const keyParts = String(columnKey).split("::");
    if (keyParts.length === 2) {
      const headerText = keyParts[1];
      selected = opts.find((o) => o.headerText === headerText) || null;
    }
  }
  
  if (!selected) throw new ValidationError("Selected column was not found in the workbook headers.");

  const task = String(taskType || "").toLowerCase();
  if (task !== "attendance" && task !== "grade") {
    throw new ValidationError("Task type must be Attendance or Grade.");
  }

  const inputList =
    task === "attendance"
      ? Array.isArray(orderedAttendanceIds) && orderedAttendanceIds.length
        ? orderedAttendanceIds.slice()
        : Array.from(attendanceIdsSet || [])
      : (gradesRows || []).map((x) => x.id);

  // preserve input file order:
  // - attendance uses Set, but we need order from txt: this preview expects caller to pass ordered list if needed.
  // For now, callers will pass gradesRows (ordered) or use parseStudentIdsText.orderedEntries to build ordered list.

  /** @type {EditorPreviewRow[]} */
  const previewRows = [];
  /** @type {ColumnMapEntry[]} */
  const columnMap = [];

  // Build per-sheet student index once
  const scopeMode = scope?.mode === "single" ? "single" : "multi";
  const selectedSheet = String(scope?.sheetName || "");
  const sheetNames =
    scopeMode === "single"
      ? [selectedSheet].filter(Boolean)
      : workbook.SheetNames.slice();

  /** @type {Map<string, { byId: Map<string, any[]>, rows: any[] }>} */
  const studentIndexBySheet = new Map();
  for (const sheetName of sheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;
    // Build student index with target IDs for better ID column detection
    const targetIdsForDetection = attendanceIdsSet || (task === "grade" && gradesRows ? new Set(gradesRows.map(r => r.id)) : null);
    studentIndexBySheet.set(sheetName, buildStudentIndexForSheet(ws, targetIdsForDetection));
  }

  // For each sheet where the selected header exists, create mapping entry and use its column for edits.
  const locations = Array.isArray(selected.locations) ? selected.locations : [];
  for (const loc of locations) {
    columnMap.push({
      sheet: loc.sheet,
      header_text: selected.headerText,
      kind: selected.kind,
      header_row: loc.header_row,
      col_letter: loc.col_letter,
    });
  }

  // Helper for old value
  function getCellDisplay(ws, addr) {
    const cell = ws?.[addr];
    if (!cell) return "";
    const v = cell.v;
    return v === null || v === undefined ? "" : v;
  }

  // Build ordered input entries: for attendance we want caller-provided ordered list; if we got Set, order is insertion.
  const orderedInput =
    task === "grade"
      ? (gradesRows || []).map((x) => ({ id: x.id, grade: x.grade }))
      : inputList.map((id) => ({ id, grade: null }));

  let idx = 0;
  for (const item of orderedInput) {
    idx += 1;
    const sid = String(item.id || "").trim();
    const desiredValue = task === "attendance" ? 1 : String(item.grade);

    // Find match across sheets where column exists (priority: same sheet order as locations)
    let matched = null;
    let matchedSheet = "";
    let ambiguous = false;
    for (const loc of locations) {
      const sheetName = loc.sheet;
      const index = studentIndexBySheet.get(sheetName);
      const hits = index?.byId?.get(sid) || [];
      if (hits.length === 1) {
        matched = { sheetName, row1: hits[0].row1, name: hits[0].name, col1: loc.col1, colLetter: loc.col_letter };
        matchedSheet = sheetName;
        break;
      }
      if (hits.length > 1) {
        ambiguous = true;
        matched = { sheetName, row1: hits[0].row1, name: hits[0].name, col1: loc.col1, colLetter: loc.col_letter };
        matchedSheet = sheetName;
        break;
      }
    }

    if (!matched) {
      previewRows.push({
        index: idx,
        input_id: sid,
        sheet: "",
        row_index1: null,
        student_id: sid,
        student_name: "",
        cell: "",
        col_letter: "",
        old_value: "",
        new_value: desiredValue,
        match_status: "notFound",
        note: "ID not found in selected scope.",
      });
      continue;
    }

    const ws = workbook.Sheets[matchedSheet];
    const addr = window.XLSX.utils.encode_cell({ r: matched.row1 - 1, c: matched.col1 - 1 });
    const oldVal = getCellDisplay(ws, addr);
    previewRows.push({
      index: idx,
      input_id: sid,
      sheet: matched.sheetName,
      row_index1: matched.row1,
      student_id: sid,
      student_name: matched.name || "",
      cell: addr,
      col_letter: matched.colLetter,
      old_value: oldVal,
      new_value: desiredValue,
      match_status: ambiguous ? "ambiguous" : "matched",
      note: ambiguous ? "Duplicate ID detected in sheet; please verify match." : "",
    });
  }

  return { preview_rows: previewRows, column_map: columnMap, selected_column: selected };
}

export function applyEditorEdits(workbook, previewRows, highlightEnabled = false, highlightColor = "#FFFF00") {
  assertXlsxLoaded();
  if (!workbook || !Array.isArray(workbook.SheetNames)) throw new ProcessingError("Invalid workbook.");
  const rows = Array.isArray(previewRows) ? previewRows : [];

  // Convert hex color to RGB format for SheetJS (strip # if present)
  let rgbColor = null;
  if (highlightEnabled && highlightColor) {
    const hex = String(highlightColor).trim();
    const cleanHex = hex.startsWith("#") ? hex.slice(1) : hex;
    // Validate hex format (6 characters, all hex digits)
    if (/^[0-9A-Fa-f]{6}$/.test(cleanHex)) {
      rgbColor = cleanHex.toUpperCase();
    } else {
      // Fall back to default yellow if invalid
      rgbColor = "FFFF00";
    }
  }

  for (const row of rows) {
    if (!row || row.match_status === "notFound") continue;
    const sheetName = String(row.sheet || "");
    const addr = String(row.cell || "");
    if (!sheetName || !addr) continue;
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const val = row.new_value;
    // Create or update cell
    const cellType =
      typeof val === "number"
        ? "n"
        : typeof val === "boolean"
          ? "b"
          : "s";
    
    // Create cell object with value
    const cell = { t: cellType, v: val };
    
    // Apply highlight color if enabled
    if (highlightEnabled && rgbColor) {
      cell.s = {
        fill: {
          patternType: "solid",
          fgColor: { rgb: rgbColor }
        }
      };
    }
    
    ws[addr] = cell;
  }

  return workbook;
}

function findHeaderColumn(ws, targetText, startCol1, endCol1, availableWeeksSet, idCol = 2) {
  // Scan rows 2-5. Stop if we hit data row (detected ID column contains ID-like number).
  const targetLower = String(targetText).toLowerCase();

  for (let r = 2; r <= 5; r++) {
    const idCheck = normalizeId(cellValue(ws, r, idCol));
    if (idCheck && /^\d+$/.test(idCheck) && idCheck.length > 5) break;

    for (let c = startCol1; c <= endCol1; c++) {
      const raw = cellValue(ws, r, c);
      const cellVal = String(raw ?? "").trim();
      if (cellVal && /^W\d+$/i.test(cellVal)) availableWeeksSet.add(cellVal.toUpperCase());
      if (cellVal.toLowerCase() === targetLower) return { headerRow: r, targetCol: c };
    }
  }
  return { headerRow: null, targetCol: null };
}

function findFirstDataRow(ws, startDataRow1, idCol = 2) {
  const range = getSheetRange(ws);
  if (!range) return startDataRow1;
  const maxRow1 = range.e.r + 1;

  const toRow = Math.min(startDataRow1 + 9, maxRow1);
  for (let r = startDataRow1; r <= toRow; r++) {
    const testId = normalizeId(cellValue(ws, r, idCol));
    if (testId && /^\d+$/.test(testId) && testId.length > 5) return r;
  }
  return startDataRow1;
}

function findLastDataRow(ws, startDataRow1, idCol = 2) {
  const range = getSheetRange(ws);
  if (!range) return startDataRow1;
  const maxRow1 = range.e.r + 1;

  for (let r = maxRow1; r >= Math.max(startDataRow1, 1); r--) {
    const testId = normalizeId(cellValue(ws, r, idCol));
    if (testId && /^\d+$/.test(testId) && testId.length > 5) return r;
  }
  return maxRow1;
}

// -----------------------------
// Attendance processing
// -----------------------------
export function processAttendance(workbook, targetIdsSet, targetWeek, attendanceType) {
  assertXlsxLoaded();
  if (!workbook || !Array.isArray(workbook.SheetNames)) throw new ProcessingError("Invalid workbook.");
  if (!(targetIdsSet instanceof Set)) throw new ProcessingError("Invalid target IDs set.");

  const foundLog = [];
  const foundIdsSet = new Set();
  const availableWeeks = new Set();

  const weekText = String(targetWeek).toUpperCase();
  const typeLower = String(attendanceType || "").toLowerCase();

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const { colSection, colLecture } = detectLectureSectionBounds(ws);
    if (!colSection) continue;

    // Detect ID and name columns dynamically for this sheet
    const { idCol, nameCol, nameCol2 } = detectIdAndNameColumns(ws, targetIdsSet);
    const finalIdCol = idCol || 2; // fallback to original assumption

    let startSearch = 0;
    let endSearch = 0;
    const range = getSheetRange(ws);
    if (!range) continue;
    const maxCol1 = range.e.c + 1;

    if (typeLower.includes("section")) {
      startSearch = colSection;
      endSearch = colLecture ? colLecture - 1 : maxCol1;
    } else if (typeLower.includes("lecture")) {
      if (!colLecture) continue;
      // Based on actual Excel structure analysis: lecture week columns (W1, W2, W3...) start
      // at the same column where "Attendance Lecture" header is found in row 1.
      // Section week columns end at colLecture - 1, so lecture starts at colLecture.
      startSearch = colLecture;
      endSearch = maxCol1;
    } else {
      throw new ValidationError("Attendance type must be 'lecture' or 'section'.");
    }

    const { headerRow, targetCol } = findHeaderColumn(ws, weekText, startSearch, endSearch, availableWeeks, finalIdCol);
    if (!targetCol) continue;

    const startDataRow = headerRow + 1;
    const firstDataRow = findFirstDataRow(ws, startDataRow, finalIdCol);
    const lastDataRow = findLastDataRow(ws, startDataRow, finalIdCol);

    // Scan rows using detected ID and name columns
    const maxRow1 = range.e.r + 1;
    for (let rowIdx = startDataRow; rowIdx <= maxRow1; rowIdx++) {
      const rawId = cellValue(ws, rowIdx, finalIdCol);
      
      // Extract ID, handling email format
      let currId = null;
      if (rawId !== null && rawId !== undefined) {
        const normalized = normalizeId(rawId);
        
        if (normalized) {
          // Check if it contains "@" (email format)
          if (normalized.includes("@")) {
            const username = normalized.split("@")[0];
            if (targetIdsSet && targetIdsSet.has(username)) {
              currId = username;
            }
          } else {
            // Direct match
            if (targetIdsSet && targetIdsSet.has(normalized)) {
              currId = normalized;
            } else if (!targetIdsSet && /^\d+$/.test(normalized) && normalized.length > 5) {
              // Pattern-based: pure numeric ID
              currId = normalized;
            }
          }
        }
      }
      
      if (!currId || !targetIdsSet.has(currId)) continue;

      foundIdsSet.add(currId);
      
      // Build name from detected name column(s)
      let nameVal = "";
      if (nameCol2 !== null) {
        // Concatenate two name columns
        const name1 = String(cellValue(ws, rowIdx, nameCol) ?? "").trim();
        const name2 = String(cellValue(ws, rowIdx, nameCol2) ?? "").trim();
        nameVal = `${name1} ${name2}`.trim();
      } else {
        // Single name column
        nameVal = String(cellValue(ws, rowIdx, nameCol) ?? "").trim();
      }
      
      const cellAddr = window.XLSX.utils.encode_cell({ r: rowIdx - 1, c: targetCol - 1 });
      const colLetter = window.XLSX.utils.encode_col(targetCol - 1);
      foundLog.push({
        sheet: sheetName,
        id: currId,
        name: nameVal || "N/A",
        target_cell: cellAddr,
        col_letter: colLetter,
        header_row: headerRow,
        first_data_row: firstDataRow,
        last_data_row: lastDataRow,
      });
    }
  }

  const notFoundIds = [];
  for (const id of targetIdsSet) if (!foundIdsSet.has(id)) notFoundIds.push(id);
  notFoundIds.sort((a, b) => String(a).localeCompare(String(b)));

  if (foundLog.length === 0 && !availableWeeks.has(weekText)) {
    const weeksStr = availableWeeks.size ? Array.from(availableWeeks).sort().join(", ") : "none found";
    throw new ProcessingError(`Week '${weekText}' was not found in any sheet. Available weeks: ${weeksStr}`);
  }

  return {
    found_log: foundLog,
    not_found_ids: notFoundIds,
    found_ids_set: foundIdsSet,
    available_weeks: Array.from(availableWeeks).sort(),
  };
}


