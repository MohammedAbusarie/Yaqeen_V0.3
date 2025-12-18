/**
 * File reading helpers (internal).
 */

/**
 * @param {File} file
 * @returns {Promise<ArrayBuffer>}
 */
export async function readFileAsArrayBuffer(file) {
  return await file.arrayBuffer();
}

/**
 * @param {File} file
 * @returns {Promise<string>}
 */
export async function readFileAsText(file) {
  return await file.text();
}


