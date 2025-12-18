/**
 * DOM helpers (internal).
 */

/**
 * @param {string} id
 * @returns {HTMLElement|null}
 */
export function $(id) {
  return document.getElementById(id);
}

/**
 * Download a string/blob content as a file.
 * @param {string} filename
 * @param {string|ArrayBuffer|Uint8Array} content
 * @param {string} mime
 */
export function downloadBlob(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}


