/**
 * Proctoring Schedule handlers.
 *
 * Single flat table with status filter dropdown.
 */

import {
  parseProctoringWorkbook,
  filterScheduleForProctor,
  getNameSuggestions,
  normalizeProctorName,
  extractDisplayName,
  areNamesSimilar,
  findSimilarNames,
} from "../proctoringParser.js";
import { fetchXlsxFromUrl, readWorkbookFromArrayBuffer } from "../../attendance.js";

const STORAGE_KEY_CUSTOM_URL = "yaqeen_proctoring_custom_url";
const STORAGE_KEY_SELECTED_NAME = "yaqeen_proctoring_selected_name";
const REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const MIN_REFRESH_INTERVAL_MS = 60 * 1000;
const STALE_THRESHOLD_MS = 10 * 60 * 1000;
const LAST_UPDATED_TICK_MS = 1000; // Update every second

/**
 * @param {{ els: any, state: import('../state.js').AppState, setStatus: (msg:string, kind?:'info'|'ok'|'error')=>void, switchView: (viewName:string)=>void, showToast?: (msg:string, opts?:any)=>void }} refs
 */
export function createProctoringHandlers(refs) {
  const { els, state, setStatus, switchView, showToast } = refs;
  
  // Track the live update timer for "Last updated" label
  let lastUpdatedTimerId = null;

  function setProctoringStatus(msg, kind = "info") {
    const statusEl = els.proctoringStatus;
    if (!statusEl) return;
    statusEl.textContent = msg || "";
    statusEl.classList.remove("is-error", "is-ok");
    if (kind === "error") statusEl.classList.add("is-error");
    if (kind === "ok") statusEl.classList.add("is-ok");
  }

  function getEffectiveUrl() {
    return state.proctoring.customUrl || state.proctoring.defaultUrl;
  }

  function updateUrlUI() {
    const urlInput = els.proctoringSheetUrl;
    const resetBtn = els.btnProctoringResetUrl;
    const urlStatus = els.proctoringUrlStatus;
    if (urlInput) urlInput.value = state.proctoring.customUrl || "";
    if (resetBtn) resetBtn.style.display = state.proctoring.customUrl ? "" : "none";
    if (urlStatus) {
      urlStatus.style.display = "";
      if (state.proctoring.customUrl) {
        urlStatus.textContent = "Using custom schedule URL";
        urlStatus.className = "proctoringUrlStatus proctoringUrlStatus--custom";
      } else {
        urlStatus.textContent = "Using default schedule URL";
        urlStatus.className = "proctoringUrlStatus proctoringUrlStatus--default";
      }
    }
  }

  function updateLastUpdatedUI() {
    const el = els.proctoringLastUpdated;
    if (!el) return;
    const last = state.proctoring.lastFetchTime;
    if (!last) {
      el.innerHTML = `<span class="lastUpdatedBadge lastUpdatedBadge--none">No schedule loaded yet</span>`;
      return;
    }
    const diffMs = Date.now() - last;
    const diffSec = Math.floor(diffMs / 1000);
    const diffMin = Math.floor(diffMs / 60000);
    let text;
    let kind = "fresh";
    
    if (diffSec < 60) {
      // Less than a minute - show seconds
      text = `${diffSec} sec ago`;
      kind = "fresh";
    } else if (diffMin < 5) {
      text = `${diffMin} min ago`;
      kind = "fresh";
    } else if (diffMin < 60) {
      text = `${diffMin} min ago`;
      kind = "recent";
    } else {
      const diffHr = Math.floor(diffMin / 60);
      text = `${diffHr}h ${diffMin % 60}m ago`;
      kind = "stale";
    }
    el.innerHTML = `<span class="lastUpdatedBadge lastUpdatedBadge--${kind}">Last updated: ${text}</span>`;
  }

  function startLastUpdatedTimer() {
    stopLastUpdatedTimer();
    lastUpdatedTimerId = setInterval(() => {
      if (state.proctoring.lastFetchTime) {
        updateLastUpdatedUI();
      }
    }, LAST_UPDATED_TICK_MS);
  }

  function stopLastUpdatedTimer() {
    if (lastUpdatedTimerId) {
      clearInterval(lastUpdatedTimerId);
      lastUpdatedTimerId = null;
    }
  }

  function updateStatsUI() {
    const el = els.proctoringStats;
    if (!el) return;
    const data = state.proctoring.parsedData;
    const selectedName = state.proctoring.selectedName;
    if (!data) { el.textContent = ""; return; }
    
    // Calculate personal stats if a name is selected
    let personalStats = "";
    if (selectedName) {
      const results = filterScheduleForProctor(data, selectedName);
      let totalPersonal = 0;
      let completedPersonal = 0;
      for (const { assignments } of results) {
        for (const a of assignments) {
          totalPersonal++;
          if (a.status === "past") completedPersonal++;
        }
      }
      if (totalPersonal > 0) {
        personalStats = ` · Your progress: ${completedPersonal}/${totalPersonal} completed`;
      }
    }
    
    el.textContent = `${data.days.length} days · ${data.totalAssignments} assignments · ${data.allProctorNames.length} proctors${personalStats}`;
  }

  function enableSearch(enabled) {
    if (els.proctoringNameSearch) els.proctoringNameSearch.disabled = !enabled;
  }

  function renderNotes(notes) {
    const container = els.proctoringNotes;
    if (!container) return;
    if (!notes || notes.length === 0) { container.style.display = "none"; container.innerHTML = ""; return; }
    container.style.display = "";
    container.innerHTML = `
      <div class="proctoringNotes__title">General Instructions</div>
      <ul class="proctoringNotes__list">${notes.map((n) => `<li>${escapeHtml(n)}</li>`).join("")}</ul>
    `;
  }

  function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * Check if a hall is a "big" hall (has groups) or "small" hall.
   * Big halls: A1-A5, B1-B5 (letter followed by single digit 1-5)
   * Small halls: A301, B501, etc. (letter followed by 3+ digits)
   * @param {string} room 
   * @returns {boolean}
   */
  function isBigHall(room) {
    if (!room) return false;
    // Match pattern: letter followed by single digit 1-5 (e.g., A1, B3)
    return /^[A-Za-z][1-5]$/.test(room.trim());
  }

  // ============================================================================
  // Loading
  // ============================================================================

  async function loadSchedule(force = false) {
    const now = Date.now();
    const last = state.proctoring.lastFetchTime;
    if (!force && last && now - last < MIN_REFRESH_INTERVAL_MS) {
      setProctoringStatus("Schedule is already up to date.", "ok");
      return;
    }

    const url = getEffectiveUrl();
    if (!url) {
      setProctoringStatus("No schedule URL configured.", "error");
      return;
    }

    state.proctoring.isLoading = true;
    state.proctoring.error = null;
    setProctoringStatus("Loading schedule…");
    enableSearch(false);
    
    // Stop the live timer during loading
    stopLastUpdatedTimer();

    // Show spinner
    const loadBtn = els.btnProctoringLoad;
    const loadText = els.proctoringLoadText;
    const loadSpinner = els.proctoringLoadSpinner;
    if (loadBtn) loadBtn.disabled = true;
    if (loadText) loadText.textContent = "Loading…";
    if (loadSpinner) loadSpinner.style.display = "inline-block";

    try {
      let arrayBuffer = null;
      try {
        arrayBuffer = await fetchXlsxFromUrl(url);
      } catch (err) {
        if (state.proctoring.workbookArrayBuffer) {
          arrayBuffer = state.proctoring.workbookArrayBuffer;
          setProctoringStatus("Using cached schedule (network update failed).", "info");
        } else {
          throw err;
        }
      }

      if (!arrayBuffer) throw new Error("Failed to download schedule.");
      state.proctoring.workbookArrayBuffer = arrayBuffer;

      const workbook = readWorkbookFromArrayBuffer(arrayBuffer);
      if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
        throw new Error("Invalid workbook.");
      }

      const parsed = parseProctoringWorkbook(workbook);
      if (!parsed || parsed.days.length === 0) {
        throw new Error("No schedule data found.");
      }

      state.proctoring.parsedData = parsed;
      state.proctoring.lastFetchTime = Date.now();
      state.proctoring.error = null;

      updateLastUpdatedUI();
      updateStatsUI();
      renderNotes(parsed.notes);
      enableSearch(true);

      if (state.proctoring.selectedName) {
        renderScheduleForName(state.proctoring.selectedName);
      } else {
        renderEmptyTable();
      }

      // Show toast instead of persistent status message
      if (showToast) {
        showToast(`Schedule loaded: ${parsed.days.length} days, ${parsed.totalAssignments} assignments.`, { kind: "ok", duration: 4000 });
      }
      setProctoringStatus("", "info");
    } catch (err) {
      state.proctoring.error = err;
      console.error("Proctoring load error:", err);
      let msg = err?.message || "Failed to load schedule.";
      if (msg.includes("CORS") || msg.includes("download")) {
        msg = "Cannot download schedule automatically. Please download the file manually from Google Sheets (File → Download → Microsoft Excel) and upload it.";
      }
      setProctoringStatus(msg, "error");
    } finally {
      state.proctoring.isLoading = false;

      // Hide spinner
      const loadBtn = els.btnProctoringLoad;
      const loadText = els.proctoringLoadText;
      const loadSpinner = els.proctoringLoadSpinner;
      if (loadBtn) loadBtn.disabled = false;
      if (loadText) loadText.textContent = "Load / Refresh";
      if (loadSpinner) loadSpinner.style.display = "none";
      
      // Restart timer if we have data (either new or cached)
      if (state.proctoring.lastFetchTime) {
        startLastUpdatedTimer();
      }
    }
  }

  // ============================================================================
  // Auto-refresh
  // ============================================================================

  function startAutoRefresh() {
    stopAutoRefresh();
    state.proctoring.autoRefreshIntervalId = setInterval(() => {
      if (document.visibilityState !== "visible") return;
      if (!state.proctoring.parsedData || !state.proctoring.lastFetchTime) return;
      if (Date.now() - state.proctoring.lastFetchTime > STALE_THRESHOLD_MS) {
        loadSchedule(false).catch(() => {});
      }
    }, REFRESH_INTERVAL_MS);
  }

  function stopAutoRefresh() {
    if (state.proctoring.autoRefreshIntervalId) {
      clearInterval(state.proctoring.autoRefreshIntervalId);
      state.proctoring.autoRefreshIntervalId = null;
    }
  }

  // ============================================================================
  // Rendering — Single Flat Table
  // ============================================================================

  function renderEmptyState(message) {
    const container = els.proctoringResults;
    if (!container) return;
    container.innerHTML = `<div class="proctoringResults__empty"><p>${escapeHtml(message)}</p></div>`;
  }

  function renderEmptyTable() {
    const container = els.proctoringResults;
    if (!container) return;
    container.innerHTML = `
      <div class="tableWrap proctoringTableWrap">
        <table class="table proctoringTable">
          <thead>
            <tr>
              <th>Date</th>
              <th>Day</th>
              <th>Period</th>
              <th>Time</th>
              <th>Room</th>
              <th>Faculty</th>
              <th>Control Room</th>
              <th>Proctors</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td colspan="9" class="proctoringResults__empty" style="text-align: center; padding: 40px;">
                <p>No records to show. Select your name to load your assignments.</p>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    `;
  }

  function getStatusPillHtml(status) {
    if (status === "past") return `<span class="statusPill statusPill--past">Passed</span>`;
    if (status === "today") return `<span class="statusPill statusPill--today">Today</span>`;
    return `<span class="statusPill statusPill--future">Upcoming</span>`;
  }

  function formatTimeRange(start, end) {
    if (!start || !end) return "";
    const fmt = (d) => {
      const h = d.getHours() % 12 || 12;
      const m = String(d.getMinutes()).padStart(2, "0");
      return `${h}:${m}`;
    };
    const period = start.getHours() < 12 ? "AM" : "PM";
    return `${fmt(start)} - ${fmt(end)} ${period}`;
  }

  function renderScheduleForName(name) {
    const container = els.proctoringResults;
    const data = state.proctoring.parsedData;
    if (!container) return;
    
    if (!data) {
      renderEmptyTable();
      return;
    }

    if (!name || !name.trim()) {
      renderEmptyTable();
      return;
    }

    const results = filterScheduleForProctor(data, name);
    if (results.length === 0) {
      container.innerHTML = `
        <div class="tableWrap proctoringTableWrap">
          <table class="table proctoringTable">
            <thead>
              <tr>
                <th>Date</th>
                <th>Day</th>
                <th>Period</th>
                <th>Time</th>
                <th>Room</th>
                <th>Faculty</th>
                <th>Control Room</th>
                <th>Proctors</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td colspan="9" class="proctoringResults__empty" style="text-align: center; padding: 40px;">
                  <p>No assignments found for "${escapeHtml(name)}".<br>Try a different spelling.</p>
                </td>
              </tr>
            </tbody>
          </table>
        </div>`;
      return;
    }
    
    // Calculate stats
    let totalCount = 0;
    let completedCount = 0;
    for (const { assignments } of results) {
      for (const a of assignments) {
        totalCount++;
        if (a.status === "past") completedCount++;
      }
    }

    const normalizedQuery = normalizeProctorName(name);
    
    // Find all similar name variations for highlighting
    const similarNames = state.proctoring.parsedData ?
      findSimilarNames(state.proctoring.parsedData.allProctorNames, name) : [];
    const normalizedSimilarNames = new Set(similarNames.map(n => normalizeProctorName(n)).filter(Boolean));

    // Flatten all rows into one array
    const allRows = [];
    for (const { day, assignments } of results) {
      for (const a of assignments) {
        allRows.push({ day, assignment: a });
      }
    }

    // Build table rows with day separators
    let lastDayKey = null;
    const rowsHtml = allRows.map(({ day, assignment }) => {
      if (!assignment || !assignment.subHalls) return "";
      const a = assignment;
      
      // Check if this is a new day (for separator)
      const currentDayKey = `${day.dateDisplay}_${day.periodEnglish}`;
      const isNewDay = lastDayKey !== null && lastDayKey !== currentDayKey;
      lastDayKey = currentDayKey;
      
      // Check if this is a big hall (to determine if we show Group #)
      const bigHall = isBigHall(a.room);
      
      const proctorHtml = a.subHalls.map((sh, idx) => {
      const proctorList = sh.proctors.map((p) => {
        const norm = normalizeProctorName(p.rawName);
        
        // Check for exact match or fuzzy match with similar names
        let isMatch = false;
        if (norm && normalizedQuery) {
          isMatch = norm === normalizedQuery || 
                    norm.includes(normalizedQuery) || 
                    normalizedQuery.includes(norm);
          
          if (!isMatch) {
            for (const similarNorm of normalizedSimilarNames) {
              if (norm === similarNorm || areNamesSimilar(p.rawName, name)) {
                isMatch = true;
                break;
              }
            }
          }
        }
        
        const display = escapeHtml(p.displayName);
        const fac = p.faculty ? ` <span class="proctorFaculty">(${escapeHtml(p.faculty)})</span>` : "";
        const cls = isMatch ? "proctorName proctorName--match" : "proctorName";
        return `<span class="${cls}">${display}${fac}</span>`;
      }).join(", ");

      // Only show Group # for big halls with multiple sub-halls
      if (a.subHalls.length > 1 && bigHall) {
        return `<div class="subHallGroup"><span class="subHallLabel">Group ${sh.groupIndex}:</span> ${proctorList}</div>`;
      }
      return proctorList;
    }).join("");

    const timeDisplay = a.startTime && a.endTime ? formatTimeRange(a.startTime, a.endTime) : escapeHtml(a.time);
    
    // Add separator row if this is a new day
    const separatorRow = isNewDay ? `<tr class="proctoringRow--daySeparator"><td colspan="9"></td></tr>` : "";

    return `${separatorRow}
      <tr class="proctoringRow proctoringRow--${a.status}" data-status="${a.status}">
        <td class="proctoringCell proctoringCell--date">${escapeHtml(day.dateDisplay)}</td>
        <td class="proctoringCell proctoringCell--day">${escapeHtml(day.dayEnglish)}</td>
        <td class="proctoringCell proctoringCell--period">${escapeHtml(day.periodEnglish)}</td>
        <td class="proctoringCell proctoringCell--time">${timeDisplay}</td>
        <td class="proctoringCell proctoringCell--room">${escapeHtml(a.room)}</td>
        <td class="proctoringCell proctoringCell--faculty">${escapeHtml(a.faculty)}</td>
        <td class="proctoringCell proctoringCell--control">${escapeHtml(a.controlRoom)}</td>
        <td class="proctoringCell proctoringCell--proctors">${proctorHtml}</td>
        <td class="proctoringCell proctoringCell--status">${getStatusPillHtml(a.status)}</td>
      </tr>
    `;
  }).join("");

  // Calculate completion percentage
  const completionPercent = totalCount > 0 ? Math.round((completedCount / totalCount) * 100) : 0;
  
  container.innerHTML = `
    <div class="proctoringAssignmentStats" style="margin-bottom: 16px; padding: 12px 16px; background: var(--surface-elevated); border-radius: var(--radius-md); border: 1px solid var(--border);">
      <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
        <div style="font-weight: 600; font-size: 15px; color: var(--text-primary);">
          ${escapeHtml(extractDisplayName(name))}'s Assignments
        </div>
        <div style="display: flex; align-items: center; gap: 12px; font-size: 14px;">
          <span style="color: var(--text-secondary);">Progress:</span>
          <span style="font-weight: 600; color: var(--success);">${completedCount}/${totalCount}</span>
          <span style="color: var(--text-muted);">(${completionPercent}%)</span>
          <div style="width: 80px; height: 6px; background: var(--surface); border-radius: 3px; overflow: hidden;">
            <div style="width: ${completionPercent}%; height: 100%; background: linear-gradient(90deg, var(--success), var(--primary)); border-radius: 3px; transition: width 0.3s ease;"></div>
          </div>
        </div>
      </div>
    </div>
    <div class="tableWrap proctoringTableWrap">
      <table class="table proctoringTable">
        <thead>
          <tr>
            <th>Date</th>
            <th>Day</th>
            <th>Period</th>
            <th>Time</th>
            <th>Room</th>
            <th>Faculty</th>
            <th>Control Room</th>
            <th>Proctors</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
  `;

  // Apply current filter
  applyStatusFilter();
}

  function applyStatusFilter() {
    const filterValue = els.proctoringStatusFilter?.value || "all";
    const rows = els.proctoringResults?.querySelectorAll?.(".proctoringRow");
    if (!rows) return;

    rows.forEach((row) => {
      const status = row.dataset.status;
      const show = filterValue === "all" || filterValue === status;
      row.style.display = show ? "" : "none";
    });
  }

  function renderSuggestions(suggestions) {
    const container = els.proctoringNameSuggestions;
    if (!container) return;
    if (!suggestions || suggestions.length === 0) {
      container.style.display = "none"; container.innerHTML = ""; return;
    }
    container.style.display = "";
    container.innerHTML = suggestions
      .map((name) => `<button type="button" class="proctoringSuggestion" data-name="${escapeHtml(name)}">${escapeHtml(name)}</button>`)
      .join("");
  }

  // ============================================================================
  // Event handlers
  // ============================================================================

  function handleProctoringLoad() {
    const inputUrl = els.proctoringSheetUrl?.value?.trim() || "";
    if (inputUrl && inputUrl !== state.proctoring.customUrl) {
      state.proctoring.customUrl = inputUrl;
      try { localStorage.setItem(STORAGE_KEY_CUSTOM_URL, inputUrl); } catch {}
      updateUrlUI();
    }
    loadSchedule(true);
  }

  function handleProctoringResetUrl() {
    state.proctoring.customUrl = "";
    try { localStorage.removeItem(STORAGE_KEY_CUSTOM_URL); } catch {}
    if (els.proctoringSheetUrl) els.proctoringSheetUrl.value = "";
    updateUrlUI();
    setProctoringStatus("Reset to default URL. Click Load to refresh.", "ok");
  }

  function handleProctoringNameInput() {
    const data = state.proctoring.parsedData;
    const query = els.proctoringNameSearch?.value || "";
    if (!data) { renderSuggestions([]); return; }
    if (!query.trim()) {
      state.proctoring.selectedName = "";
      try { localStorage.removeItem(STORAGE_KEY_SELECTED_NAME); } catch {}
      renderSuggestions([]);
      renderEmptyTable();
      updateStatsUI();
      return;
    }
    renderSuggestions(getNameSuggestions(data, query, 10));
  }

  async function handleProctoringSuggestionClick(e) {
    const btn = e.target.closest(".proctoringSuggestion");
    if (!btn) return;
    const name = btn.dataset.name;
    if (!name) return;
    
    // Save selected name to state and localStorage
    if (els.proctoringNameSearch) els.proctoringNameSearch.value = name;
    state.proctoring.selectedName = name;
    try { localStorage.setItem(STORAGE_KEY_SELECTED_NAME, name); } catch {}
    
    renderSuggestions([]);
    
    // Auto-load schedule if not already loaded, then render
    if (!state.proctoring.parsedData) {
      await loadSchedule(false);
    }
    renderScheduleForName(name);
    updateStatsUI();
  }

  function handleProctoringNameKeydown(e) {
    if (e.key === "Enter") {
      const query = els.proctoringNameSearch?.value || "";
      state.proctoring.selectedName = query;
      try { localStorage.setItem(STORAGE_KEY_SELECTED_NAME, query); } catch {}
      renderSuggestions([]);
      renderScheduleForName(query);
      updateStatsUI();
    }
  }

  function handleStatusFilterChange() {
    applyStatusFilter();
  }

  function handleToggleStrikethrough() {
    const rows = els.proctoringResults?.querySelectorAll?.(".proctoringRow--past");
    if (!rows) return;
    
    const icon = els.strikethroughIcon;
    let isEnabled = icon?.textContent === "✓";
    
    rows.forEach((row) => {
      if (isEnabled) {
        row.classList.add("no-dim");
      } else {
        row.classList.remove("no-dim");
      }
    });
    
    if (icon) icon.textContent = isEnabled ? "✗" : "✓";
    
    // Update button text
    const btn = els.btnToggleStrikethrough;
    if (btn) {
      btn.innerHTML = `<span id="strikethroughIcon">${isEnabled ? "✗" : "✓"}</span> Dim Passed`;
    }
  }

  function handleVisibilityChange() {
    if (document.visibilityState === "visible") {
      const last = state.proctoring.lastFetchTime;
      if (last && Date.now() - last > STALE_THRESHOLD_MS && state.proctoring.parsedData) {
        loadSchedule(false).catch(() => {});
      }
    }
  }

  // ============================================================================
  // Init
  // ============================================================================

  function initProctoring() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_CUSTOM_URL);
      if (saved) state.proctoring.customUrl = saved;
    } catch {}
    
    // Load saved name from localStorage
    try {
      const savedName = localStorage.getItem(STORAGE_KEY_SELECTED_NAME);
      if (savedName) {
        state.proctoring.selectedName = savedName;
        if (els.proctoringNameSearch) els.proctoringNameSearch.value = savedName;
      }
    } catch {}
    
    updateUrlUI();
    
    // Render empty table initially
    renderEmptyTable();
    
    // If there's already a last fetch time (cached data), start the live timer
    if (state.proctoring.lastFetchTime) {
      updateLastUpdatedUI();
      startLastUpdatedTimer();
    }
    
    startAutoRefresh();
    document.removeEventListener("visibilitychange", handleVisibilityChange);
    document.addEventListener("visibilitychange", handleVisibilityChange);
  }

  return {
    handleProctoringLoad,
    handleProctoringResetUrl,
    handleProctoringNameInput,
    handleProctoringSuggestionClick,
    handleProctoringNameKeydown,
    handleStatusFilterChange,
    handleToggleStrikethrough,
    initProctoring,
    loadSchedule,
  };
}
