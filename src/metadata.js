/**
 * Metadata + naming helpers (internal).
 */

/**
 * @param {string} name
 * @returns {string}
 */
export function safeBaseName(name) {
  return String(name)
    .replace(/\.[^.]+$/, "")
    .replace(/[^\w\-]+/g, "_")
    .slice(0, 60);
}

/**
 * @param {Object} args
 * @param {number} args.weekNum
 * @param {'lecture'|'section'|string} args.type
 * @param {string} args.sheetUrl
 * @param {number} args.totalLoadedUnique
 * @param {number} args.orderedEntriesLen
 * @param {number} args.foundCount
 * @param {number} args.notFoundCount
 */
export function buildMetadata({
  weekNum,
  type,
  sheetUrl,
  totalLoadedUnique,
  orderedEntriesLen,
  foundCount,
  notFoundCount,
}) {
  const ts = new Date();
  const pad2 = (n) => String(n).padStart(2, "0");
  const timestamp = `${ts.getFullYear()}-${pad2(ts.getMonth() + 1)}-${pad2(ts.getDate())} ${pad2(
    ts.getHours()
  )}:${pad2(ts.getMinutes())}:${pad2(ts.getSeconds())}`;

  return {
    timestamp,
    week: `W${weekNum}`,
    type: type === "lecture" ? "Lecture" : "Section",
    sheet_url: sheetUrl || "",
    total_loaded: totalLoadedUnique,
    total_input_rows: orderedEntriesLen,
    total_found: foundCount,
    total_not_found: notFoundCount,
  };
}


