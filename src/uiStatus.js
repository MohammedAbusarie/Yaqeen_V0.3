/**
 * Status + loading UI helpers (internal).
 */

/**
 * @param {{status?: HTMLElement}} els
 * @param {string} msg
 * @param {'info'|'ok'|'error'} [kind]
 */
export function setStatus(els, msg, kind = "info") {
  if (els.status) {
    els.status.textContent = msg || "";
    els.status.classList.toggle("is-error", kind === "error");
    els.status.classList.toggle("is-ok", kind === "ok");
  }
}

/**
 * @param {{btnRun?: HTMLButtonElement}} els
 * @param {boolean} disabled
 * @param {string|null} [label]
 */
export function disableRun(els, disabled, label = null) {
  // This function is kept for compatibility but no longer used
  // Editor uses its own status/loading mechanisms
  if (els.btnRun) {
    els.btnRun.disabled = disabled;
    els.btnRun.classList.toggle("btn--loading", disabled);
    if (label !== null) {
      const textEl = els.btnRun.querySelector(".btn__text");
      if (textEl) textEl.textContent = label;
    }
  }
}


