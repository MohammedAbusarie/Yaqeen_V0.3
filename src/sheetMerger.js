// Core sheet merger logic
// Handles workbook parsing, column detection, header scanning, and merging operations

import { ValidationError, FileError, readWorkbookFromArrayBuffer } from "../attendance.js";

/**
 * @typedef {Object} ColumnInfo
 * @property {string} sheet - Sheet name
 * @property {number} columnIndex - 0-based column index
 * @property {string} headerText - Detected header text
 * @property {number} headerRow - 1-based row number where header was found
 * @property {string} columnLetter - Excel-style column letter (A, B, C, etc.)
 * @property {Array<string>} sampleValues - Sample data values from rows below header
 * @property {string} key - Unique key: `${sheet}::${columnLetter}`
 */

/**
 * @typedef {Object} MergedData
 * @property {Array<Array<any>>} rows - All rows (header + data)
 * @property {Array<string>} headers - Column headers
 * @property {number} totalRows - Total row count
 * @property {number} totalColumns - Total column count
 * @property {Array<string>} sourceSheets - List of source sheet names
 */

/**
 * Convert 0-based column index to Excel-style letter (0=A, 1=B, 25=Z, 26=AA, etc.)
 */
export function columnIndexToLetter(index) {
  let letter = "";
  while (index >= 0) {
    letter = String.fromCharCode((index % 26) + 65) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

/**
 * Scan rows 1-5 across all sheets to detect potential column headers.
 * Returns a map of unique header strings found and their locations.
 * 
 * @param {any} workbook - SheetJS workbook object
 * @param {Array<number>} scanRows - 1-based row numbers to scan (default: [1,2,3,4,5])
 * @returns {Record<string, Array<{sheet: string, row: number, col: number}>>}
 */
export function scanHeadersFromRows(workbook, scanRows = [1, 2, 3, 4, 5]) {
  if (!workbook || !workbook.SheetNames) {
    throw new ValidationError("Invalid workbook");
  }

  const headerMap = {}; // headerText -> [{sheet, row, col}, ...]

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    for (const rowNum of scanRows) {
      const range = window.XLSX.utils.decode_range(sheet["!ref"] || "A1");
      for (let colIdx = range.s.c; colIdx <= range.e.c; colIdx++) {
        const cellRef = window.XLSX.utils.encode_cell({ r: rowNum - 1, c: colIdx });
        const cell = sheet[cellRef];
        if (cell && cell.v && typeof cell.v === "string" && cell.v.trim()) {
          const headerText = String(cell.v).trim();
          if (!headerMap[headerText]) {
            headerMap[headerText] = [];
          }
          headerMap[headerText].push({
            sheet: sheetName,
            row: rowNum,
            col: colIdx,
          });
        }
      }
    }
  }

  return headerMap;
}

/**
 * Extract all columns from all sheets with header detection and sample values.
 * Scans rows 1-5 for headers, then extracts up to 5 sample data values.
 * 
 * @param {any} workbook - SheetJS workbook object
 * @returns {Array<ColumnInfo>}
 */
export function extractAllColumns(workbook) {
  if (!workbook || !workbook.SheetNames) {
    throw new ValidationError("Invalid workbook");
  }

  const columns = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet || !sheet["!ref"]) continue;

    const range = window.XLSX.utils.decode_range(sheet["!ref"]);
    const maxCol = range.e.c;
    const maxRow = range.e.r;

    for (let colIdx = 0; colIdx <= maxCol; colIdx++) {
      const columnLetter = columnIndexToLetter(colIdx);
      
      // Scan rows 1-5 for header
      let headerText = `Column ${columnLetter}`;
      let headerRow = 1;
      
      for (let rowIdx = 0; rowIdx < Math.min(5, maxRow + 1); rowIdx++) {
        const cellRef = window.XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
        const cell = sheet[cellRef];
        if (cell && cell.v && typeof cell.v === "string" && cell.v.trim()) {
          headerText = String(cell.v).trim();
          headerRow = rowIdx + 1;
          break;
        }
      }

      // Extract sample values (up to 5 rows after header)
      const sampleValues = [];
      const startRow = headerRow; // Start from row after header
      for (let rowIdx = startRow; rowIdx <= Math.min(startRow + 4, maxRow); rowIdx++) {
        const cellRef = window.XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
        const cell = sheet[cellRef];
        if (cell && cell.v !== undefined && cell.v !== null) {
          const value = String(cell.v).trim();
          if (value) {
            sampleValues.push(value);
          }
        }
      }

      columns.push({
        sheet: sheetName,
        columnIndex: colIdx,
        headerText,
        headerRow,
        columnLetter,
        sampleValues,
        key: `${sheetName}::${columnLetter}`,
      });
    }
  }

  return columns;
}

/**
 * Build an empty mapping matrix structure.
 * 
 * @param {Array<string>} sheetNames - List of sheet names
 * @param {number} maxPositions - Maximum number of output column positions
 * @returns {Record<string, Record<number, string|null>>} - {sheetName: {position: columnKey|null}}
 */
export function buildMappingMatrix(sheetNames, maxPositions = 10) {
  const matrix = {};
  for (const sheetName of sheetNames) {
    matrix[sheetName] = {};
    for (let pos = 0; pos < maxPositions; pos++) {
      matrix[sheetName][pos] = null;
    }
  }
  return matrix;
}

/**
 * Merge columns sequentially based on mapping.
 * For each output position, collect mapped columns from each sheet and concatenate rows sequentially.
 * 
 * @param {any} workbook - SheetJS workbook object
 * @param {Record<string, Record<number, string>>} mapping - {sheetName: {position: columnKey}}
 * @param {Array<ColumnInfo>} allColumns - All available columns
 * @param {boolean} eliminateHeaders - Whether to remove duplicate header rows
 * @returns {MergedData}
 */
export function mergeColumnsSequentially(workbook, mapping, allColumns, eliminateHeaders = false) {
  if (!workbook || !mapping) {
    throw new ValidationError("Invalid workbook or mapping");
  }

  // Determine number of output columns
  let maxPosition = -1;
  for (const sheetName in mapping) {
    for (const posStr in mapping[sheetName]) {
      const pos = parseInt(posStr, 10);
      if (pos > maxPosition) maxPosition = pos;
    }
  }

  if (maxPosition < 0) {
    throw new ValidationError("No columns mapped. Please drag at least one column to the matrix.");
  }

  const numOutputCols = maxPosition + 1;
  
  // Build column info lookup
  const columnLookup = {};
  for (const col of allColumns) {
    columnLookup[col.key] = col;
  }

  // Collect headers for output columns (from first mapped column in each position)
  const headers = [];
  for (let pos = 0; pos < numOutputCols; pos++) {
    let headerFound = false;
    for (const sheetName of workbook.SheetNames) {
      const columnKey = mapping[sheetName]?.[pos];
      if (columnKey && columnLookup[columnKey]) {
        headers.push(columnLookup[columnKey].headerText);
        headerFound = true;
        break;
      }
    }
    if (!headerFound) {
      headers.push(`Column ${pos + 1}`);
    }
  }

  // Merge rows
  const mergedRows = [];
  const sourceSheets = new Set();

  // When eliminating headers: add a single header row at the top (data rows will skip their headers)
  // When NOT eliminating: headers will come from each sheet's data rows
  if (eliminateHeaders) {
    mergedRows.push(headers);
  }

  // For each sheet, extract and concatenate rows sequentially
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet || !sheet["!ref"]) continue;

    const range = window.XLSX.utils.decode_range(sheet["!ref"]);
    const maxRow = range.e.r;

    // Find the minimum header row across mapped columns for this sheet
    let minHeaderRow = Infinity;
    for (let pos = 0; pos < numOutputCols; pos++) {
      const columnKey = mapping[sheetName]?.[pos];
      if (columnKey && columnLookup[columnKey]) {
        const headerRow = columnLookup[columnKey].headerRow;
        if (headerRow < minHeaderRow) {
          minHeaderRow = headerRow;
        }
      }
    }

    // If no columns mapped for this sheet, skip it
    if (minHeaderRow === Infinity) {
      continue;
    }

    sourceSheets.add(sheetName);

    // minHeaderRow is 1-based. In 0-based indexing:
    // - (minHeaderRow - 1) = the header row itself
    // - minHeaderRow = the row AFTER the header
    // When eliminating: skip header by starting from minHeaderRow
    // When NOT eliminating: include header by starting from (minHeaderRow - 1)
    const startDataRow = eliminateHeaders ? minHeaderRow : (minHeaderRow - 1);
    
    for (let rowIdx = startDataRow; rowIdx <= maxRow; rowIdx++) {
      const row = [];
      let hasData = false;

      for (let pos = 0; pos < numOutputCols; pos++) {
        const columnKey = mapping[sheetName]?.[pos];
        let cellValue = "";

        if (columnKey && columnLookup[columnKey]) {
          const col = columnLookup[columnKey];
          const cellRef = window.XLSX.utils.encode_cell({ r: rowIdx, c: col.columnIndex });
          const cell = sheet[cellRef];
          
          if (cell && cell.v !== undefined && cell.v !== null) {
            cellValue = cell.v;
            hasData = true;
          }
        }

        row.push(cellValue);
      }

      // Only add row if it has at least some data
      if (hasData) {
        mergedRows.push(row);
      }
    }
  }

  // Fallback: ensure at least the header exists if no data rows
  if (mergedRows.length === 0) {
    mergedRows.push(headers);
  }

  return {
    rows: mergedRows,
    headers,
    totalRows: mergedRows.length,
    totalColumns: numOutputCols,
    sourceSheets: Array.from(sourceSheets),
  };
}

/**
 * Generate a new XLSX workbook from merged data.
 * 
 * @param {MergedData} mergedData - Merged data structure
 * @param {string} outputSheetName - Name for the output sheet
 * @returns {any} - SheetJS workbook object
 */
export function generateMergedWorkbook(mergedData, outputSheetName = "Merged") {
  if (!mergedData || !mergedData.rows || mergedData.rows.length === 0) {
    throw new ValidationError("No data to export");
  }

  const wb = window.XLSX.utils.book_new();
  const ws = window.XLSX.utils.aoa_to_sheet(mergedData.rows);
  window.XLSX.utils.book_append_sheet(wb, ws, outputSheetName);

  return wb;
}

