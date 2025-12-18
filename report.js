// Report generation (ported from gui/core/attendance.py generate_report_data/generate_text_report)
import { normalizeId } from "./attendance.js";

/**
 * @typedef {Object} ReportMetadata
 * @property {string} [timestamp]
 * @property {string} [week]
 * @property {string} [type]
 * @property {string} [sheet_url]
 * @property {number} [total_loaded]
 * @property {number} [total_input_rows]
 * @property {number} [total_found]
 * @property {number} [total_not_found]
 */

/**
 * @typedef {Object} ReportStudentRow
 * @property {string} id
 * @property {number} count
 * @property {string} name
 * @property {string} cell
 */

/**
 * @typedef {Object} ReportSheet
 * @property {string} name
 * @property {string} column
 * @property {number} header_row
 * @property {ReportStudentRow[]} students
 */

/**
 * @typedef {Object} OrderedReportRow
 * @property {number} index
 * @property {'title'|'id'} type
 * @property {string} [title]
 * @property {string} id
 * @property {string|number} count
 * @property {string} name
 * @property {string} sheet
 * @property {string} cell
 */

/**
 * @typedef {Object} ReportData
 * @property {ReportMetadata} metadata
 * @property {ReportSheet[]} sheets
 * @property {string[]} not_found
 * @property {unknown} [input_order]
 * @property {OrderedReportRow[]} [ordered_rows]
 * @property {Record<string, number>} [id_counts]
 */

export function generateReportData({
  foundLog,
  notFoundIds,
  metadata,
  orderedEntries = null,
  idCounts = null,
  sectionIdCounts = null,
}) {
  const idCountsSafe = idCounts || {};
  const sectionIdCountsSafe = sectionIdCounts || {};

  // --------
  // Sheets
  // --------
  const sheetNames = Array.from(new Set((foundLog || []).map((x) => x.sheet))).sort((a, b) =>
    String(a).localeCompare(String(b))
  );

  const sheets = sheetNames.map((sheet) => {
    const sheetHits = (foundLog || []).filter((i) => i.sheet === sheet);
    const colLetter = sheetHits[0]?.col_letter || "";
    const headerRow = sheetHits[0]?.header_row ?? 2;

    return {
      name: sheet,
      column: colLetter,
      header_row: headerRow,
      students: sheetHits.map((hit) => ({
        id: hit.id,
        count: Number.parseInt(idCountsSafe[hit.id] ?? 1, 10),
        name: hit.name ?? "N/A",
        cell: hit.target_cell,
      })),
    };
  });

  const report = {
    metadata: metadata || {},
    sheets,
    not_found: notFoundIds || [],
  };

  // -----------------------------
  // Optional ordered input section (matches GUI JSON keys)
  // -----------------------------
  // Optional ordered input section (matches GUI JSON keys)
  if (orderedEntries !== null && orderedEntries !== undefined) {
    const foundById = {};
    for (const hit of foundLog || []) {
      const sid = hit?.id;
      if (sid && !foundById[sid]) foundById[sid] = hit;
    }

    const orderedRows = [];
    let idx = 0;
    for (const entry of orderedEntries || []) {
      idx += 1;
      if (entry && typeof entry === "object" && entry.type === "title") {
        const title = String(entry.title || "").trim();
        orderedRows.push({
          index: idx,
          type: "title",
          title,
          id: "",
          count: "",
          name: title,
          sheet: "",
          cell: "",
        });
        continue;
      }

      const sid = normalizeId(entry?.id ?? entry);
      const hit = sid ? foundById[sid] : null;
      const sectionId = entry?.section ?? null;
      let sectionCount = "";
      if (sectionId !== null && sectionId !== undefined) {
        sectionCount = Number.parseInt(sectionIdCountsSafe?.[sectionId]?.[sid] ?? 0, 10);
      }

      orderedRows.push({
        index: idx,
        type: "id",
        id: sid,
        count: sectionId !== null ? sectionCount : "",
        name: (hit?.name ?? "") || "",
        sheet: (hit?.sheet ?? "") || "",
        cell: (hit?.target_cell ?? "") || "",
      });
    }

    report.input_order = orderedEntries;
    report.ordered_rows = orderedRows;
    report.id_counts = idCountsSafe;
  }

  return report;
}

export function generateTextReport(reportData) {
  const metadata = reportData?.metadata || {};
  const week = metadata.week || "";
  const type = metadata.type || "";
  const timestamp = metadata.timestamp || "";

  const lines = [];
  lines.push("ATTENDANCE ACTION REPORT");
  lines.push("========================");
  lines.push(`Target: ${week} | Type: ${type}`);
  lines.push(`Generated: ${timestamp}`);
  lines.push("");

  const sheets = Array.isArray(reportData?.sheets) ? reportData.sheets : [];
  for (const sheetData of sheets) {
    lines.push(`📂 SHEET: ${sheetData.name}`);
    lines.push(
      `   Column to Mark: ${sheetData.column} (Header: ${week} on Row ${sheetData.header_row})`
    );
    lines.push("-".padEnd(60, "-"));
    lines.push("   ACTION: Go to these cells and type '1'");
    lines.push("-".padEnd(60, "-"));

    for (const student of sheetData.students || []) {
      const count = student?.count ?? 1;
      const cell = String(student?.cell ?? "").padEnd(5, " ");
      lines.push(
        `   [ ] Cell ${cell} <- ${student.id} (x${count}) (${student.name ?? "N/A"})`
      );
    }
    lines.push("");
    lines.push("=".padEnd(60, "="));
    lines.push("");
  }

  lines.push("❌ NOT FOUND STUDENTS");
  lines.push("=====================");
  const notFound = Array.isArray(reportData?.not_found) ? reportData.not_found : [];
  if (notFound.length) {
    for (const nid of notFound) lines.push(`   ID: ${nid}`);
  } else {
    lines.push("   (All students found!)");
  }

  return lines.join("\n");
}


