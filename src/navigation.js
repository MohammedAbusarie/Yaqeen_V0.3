/**
 * Navigation/view switching (internal).
 */

/**
 * @param {(id: string) => (HTMLElement|null)} getEl
 * @param {'home'|'inputs'|'report'|'ocr'|'sheetMerger'|'qrTool'|'about'} viewName
 */
export function switchView(getEl, viewName) {
  const views = ["home", "inputs", "report", "ocr", "sheetMerger", "qrTool", "about"];
  views.forEach((v) => {
    const id = v === "qrTool" ? "viewQrTool" : `view${v.charAt(0).toUpperCase() + v.slice(1)}`;
    const viewEl = getEl(id);
    const navBtn = getEl(`nav${v.charAt(0).toUpperCase() + v.slice(1)}`);
    if (viewEl) {
      if (v === viewName) {
        viewEl.classList.add("view--active");
        viewEl.style.display = "";
        if (navBtn) navBtn.classList.add("is-active");
      } else {
        viewEl.classList.remove("view--active");
        if (v === "home" || v === "ocr" || v === "sheetMerger" || v === "qrTool") {
          viewEl.style.display = "none";
        }
        if (navBtn) navBtn.classList.remove("is-active");
      }
    }
  });
  const footerTip = getEl("footerTip");
  if (footerTip) {
    footerTip.style.display = viewName === "inputs" ? "block" : "none";
  }
}


