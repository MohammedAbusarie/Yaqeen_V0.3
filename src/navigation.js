/**
 * Navigation/view switching (internal).
 */

/**
 * @param {(id: string) => (HTMLElement|null)} getEl
 * @param {'inputs'|'report'|'about'|string} viewName
 */
export function switchView(getEl, viewName) {
  const views = ["inputs", "report", "about"];
  views.forEach((v) => {
    const viewEl = getEl(`view${v.charAt(0).toUpperCase() + v.slice(1)}`);
    const navBtn = getEl(`nav${v.charAt(0).toUpperCase() + v.slice(1)}`);
    if (viewEl) {
      if (v === viewName) {
        viewEl.classList.add("view--active");
        if (navBtn) navBtn.classList.add("is-active");
      } else {
        viewEl.classList.remove("view--active");
        if (navBtn) navBtn.classList.remove("is-active");
      }
    }
  });
}


