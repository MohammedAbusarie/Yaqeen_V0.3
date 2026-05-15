/**
 * Proctoring schedule parser.
 *
 * Pure functions that transform a SheetJS workbook into structured,
 * normalized schedule data. Extremely defensive against human-edited
 * spreadsheets: inconsistent formatting, typos, missing values, merged
 * cells, newline-separated names, space-separated names, variable column
 * counts, and mixed faculty suffixes.
 */

/**
 * @typedef {Object} ProctorEntry
 * @property {string} rawName
 * @property {string} displayName
 * @property {string|null} faculty
 * @property {boolean} hasHonorific
 */

/**
 * @typedef {Object} SubHall
 * @property {number} groupIndex
 * @property {ProctorEntry[]} proctors
 */

/**
 * @typedef {Object} ScheduleAssignment
 * @property {string} room
 * @property {string} time
 * @property {string} faculty
 * @property {string} controlRoom
 * @property {SubHall[]} subHalls
 * @property {number} rawRow
 * @property {boolean} incomplete
 * @property {Date|null} startTime
 * @property {Date|null} endTime
 * @property {'past'|'today'|'future'} status
 */

/**
 * @typedef {Object} DaySchedule
 * @property {string} dayArabic
 * @property {string} dayEnglish
 * @property {string} periodArabic
 * @property {string} periodEnglish
 * @property {Date|null} date
 * @property {string} dateDisplay
 * @property {string} originalName
 * @property {ScheduleAssignment[]} assignments
 * @property {'past'|'today'|'future'} dayStatus
 */

/**
 * @typedef {Object} ParsedProctoringData
 * @property {DaySchedule[]} days
 * @property {Map<string, string>} controlRooms
 * @property {string[]} allProctorNames
 * @property {string[]} notes
 * @property {number} totalAssignments
 */

const UNSPECIFIED = "[Unspecified]";

const DAY_MAP = {
  "السبت": "Saturday",
  "الاحد": "Sunday",
  "الاثنين": "Monday",
  "الثلاثاء": "Tuesday",
  "الاربعاء": "Wednesday",
  "الخميس": "Thursday",
  "الجمعة": "Friday",
};

const PERIOD_MAP = {
  "اولي": "Period 1",
  "تانيه": "Period 2",
  "تالته": "Period 3",
  "رابعه": "Period 4",
};

/**
 * Parse a number from a sheet name into a calendar date.
 * The convention is DDMM without separator.
 *
 * @param {number} num
 * @returns {{day:number, month:number}|null}
 */
export function parseDateNumber(num) {
  if (!Number.isFinite(num) || num < 1) return null;
  const str = String(Math.floor(num));

  if (str.length >= 3) {
    const month2 = parseInt(str.slice(-2), 10);
    if (month2 >= 10 && month2 <= 12) {
      const day = parseInt(str.slice(0, -2), 10);
      if (day >= 1 && day <= 31) return { day, month: month2 };
    }
  }

  if (str.length >= 2) {
    const month1 = parseInt(str.slice(-1), 10);
    const day = parseInt(str.slice(0, -1), 10);
    if (day >= 1 && day <= 31 && month1 >= 1 && month1 <= 9) return { day, month: month1 };
  }

  if (str.length === 2) {
    const day = parseInt(str[0], 10);
    const month = parseInt(str[1], 10);
    if (day >= 1 && day <= 9 && month >= 1 && month <= 9) return { day, month };
  }

  return null;
}

/**
 * Extract day, date number, and period from a sheet name.
 * Handles common typos like "فتزه" instead of "فتره" and missing spaces.
 *
 * @param {string} sheetName
 * @returns {{day:string, dateNum:number|null, period:string, originalName:string}|null}
 */
export function parseSheetName(sheetName) {
  if (!sheetName || typeof sheetName !== "string") return null;
  const normalized = sheetName
    .replace(/فتزه/g, "فتره")
    .replace(/\s+/g, " ")
    .trim();
  const regex = /^(السبت|الاحد|الاثنين|الثلاثاء|الاربعاء|الخميس|الجمعة)\s*(\d+)\s+فتره\s+(اولي|تانيه|تالته|رابعه)$/;
  const match = normalized.match(regex);
  if (!match) return null;
  return { day: match[1], dateNum: parseInt(match[2], 10), period: match[3], originalName: sheetName };
}

function buildDateFromParsed(parsed) {
  if (!parsed || parsed.dateNum == null) return null;
  const dm = parseDateNumber(parsed.dateNum);
  if (!dm) return null;
  const year = new Date().getFullYear();
  const date = new Date(year, dm.month - 1, dm.day);
  if (date.getMonth() !== dm.month - 1 || date.getDate() !== dm.day) return null;
  return date;
}

function formatDateDisplay(date) {
  if (!date) return UNSPECIFIED;
  try {
    return date.toLocaleDateString("en-GB", {
      weekday: "long",
      year: "numeric",
      month: "short",
      day: "numeric",
    });
  } catch {
    return `${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear()}`;
  }
}

/**
 * Normalize an Arabic/English name for robust fuzzy matching.
 *
 * @param {string} raw
 * @returns {string|null}
 */
export function normalizeProctorName(raw) {
  if (!raw || typeof raw !== "string") return null;
  let s = raw.replace(/\n/g, " ").replace(/\s+/g, " ").trim();
  if (!s) return null;
  if (/^=/.test(s)) return null;
  if (/^\d+(\.\d+)?$/.test(s)) return null;
  if (s.length < 2) return null;
  s = s.replace(/\s*\(\s*[A-Z]+\s*\)\s*(?:م)?\s*$/i, "").trim();
  if (!s || s.length < 2) return null;
  s = s.replace(/[\u064B-\u065F\u0670]/g, "");
  s = s.replace(/[إأآا]/g, "ا");
  s = s.replace(/\s+/g, " ").trim();
  return s.toLowerCase();
}

/**
 * Calculate Levenshtein distance between two strings.
 * This measures how many single-character edits are needed to change one string into another.
 *
 * @param {string} a
 * @param {string} b
 * @returns {number}
 */
function levenshteinDistance(a, b) {
  const matrix = [];
  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

/**
 * Check if two names are similar enough to be considered typos of the same person.
 *
 * Strategy: compare word-by-word rather than whole-string Levenshtein.
 * Each corresponding word must differ by at most 1 character edit.
 * A whole-word difference (distance ≥ 2) means it's a different word, not a typo.
 *
 * Examples:
 *   "محمد حاتم" vs "محمد حتم"   → word[1]: ح=ح, dist=1  → MATCH   (typo: missing ا)
 *   "فرح حمدي"  vs "فرح محمد"  → word[1]: ح≠م         → NO MATCH (different first letter)
 *   "محمد ايهاب" vs "احمد ايهاب" → word[0]: م≠ا        → NO MATCH (different first letter)
 *
 * @param {string} name1
 * @param {string} name2
 * @returns {boolean}
 */
export function areNamesSimilar(name1, name2) {
  if (!name1 || !name2) return false;

  const norm1 = normalizeProctorName(name1);
  const norm2 = normalizeProctorName(name2);

  if (!norm1 || !norm2) return false;
  if (norm1 === norm2) return true;

  const words1 = norm1.split(/\s+/);
  const words2 = norm2.split(/\s+/);

  // Names with very different word counts are structurally different people.
  if (Math.abs(words1.length - words2.length) > 1) return false;

  if (words1.length === words2.length) {
    for (let i = 0; i < words1.length; i++) {
      const w1 = words1[i];
      const w2 = words2[i];
      // Different first letter = different name, not a typo (محمد ≠ احمد).
      if (w1[0] !== w2[0]) return false;
      if (levenshteinDistance(w1, w2) > 1) return false;
    }
    return true;
  }

  // Word counts differ by exactly 1: try each possible word-skip alignment.
  const [shorter, longer] = words1.length < words2.length
    ? [words1, words2]
    : [words2, words1];

  for (let skip = 0; skip < longer.length; skip++) {
    const aligned = longer.filter((_, i) => i !== skip);
    const allClose = aligned.every((w, i) => {
      const s = shorter[i];
      return w[0] === s[0] && levenshteinDistance(w, s) <= 1;
    });
    if (allClose) return true;
  }

  return false;
}

/**
 * Find all names similar to the query from a list of names.
 * Returns exact/substring matches and word-level typo variants.
 *
 * @param {string[]} allNames
 * @param {string} query
 * @returns {string[]}
 */
export function findSimilarNames(allNames, query) {
  const normalizedQuery = normalizeProctorName(query);
  if (!normalizedQuery) return [];

  const similarNames = [];
  const seen = new Set();

  for (const name of allNames) {
    if (seen.has(name)) continue;

    const normalizedName = normalizeProctorName(name);
    if (!normalizedName) continue;

    const isMatch =
      normalizedName === normalizedQuery ||
      normalizedName.includes(normalizedQuery) ||
      normalizedQuery.includes(normalizedName) ||
      areNamesSimilar(name, query);

    if (isMatch) {
      similarNames.push(name);
      seen.add(name);
    }
  }

  return similarNames;
}

/**
 * Extract a clean display name from a raw cell value.
 *
 * @param {string} raw
 * @returns {string}
 */
export function extractDisplayName(raw) {
  if (!raw || typeof raw !== "string") return UNSPECIFIED;
  const cleaned = raw.replace(/\n/g, " ").replace(/\s+/g, " ").trim();
  if (!cleaned) return UNSPECIFIED;
  const withoutSuffix = cleaned.replace(/\s*\(\s*[A-Z]+\s*\)\s*(?:م)?\s*$/i, "").trim();
  return withoutSuffix || cleaned;
}

function cleanCell(cell) {
  if (cell == null) return null;
  const str = String(cell).replace(/\s+/g, " ").trim();
  return str || null;
}

/**
 * Split a proctor cell into individual proctor entries.
 * Handles both newline-separated and space-separated names with faculty suffixes.
 *
 * Examples:
 *   "محمد ابوسريع (CS) امل معوض" → 2 entries
 *   "نوران قدري (ART)\nمنى ناجى (H)" → 2 entries
 *   "حبيبة اشرف ( MEDIA)" → 1 entry
 *   "تقى احمد" → 1 entry (no faculty)
 *
 * @param {string} cellText
 * @returns {ProctorEntry[]}
 */
export function splitProctorCell(cellText) {
  const text = String(cellText || "").trim();
  if (!text) return [];

  const lines = text.split("\n").map((s) => s.trim()).filter(Boolean);
  const results = [];

  const facultyPattern = /([^\n()]+?)\s*\(\s*([A-Z]+)\s*\)(?:\s*م)?/gi;

  for (const line of lines) {
    const matches = Array.from(line.matchAll(facultyPattern));

    if (matches.length === 0) {
      // No faculty suffix found - treat whole line as one name
      results.push({
        rawName: line,
        displayName: extractDisplayName(line),
        faculty: null,
        hasHonorific: false,
      });
      continue;
    }

    let lastEnd = 0;
    for (const match of matches) {
      const start = match.index;
      const end = start + match[0].length;

      const between = line.slice(lastEnd, start).trim();
      if (between.length > 2) {
        results.push({
          rawName: between,
          displayName: extractDisplayName(between),
          faculty: null,
          hasHonorific: false,
        });
      }

      results.push({
        rawName: match[0].trim(),
        displayName: extractDisplayName(match[1]),
        faculty: match[2].toUpperCase(),
        hasHonorific: /م\s*$/.test(match[0]),
      });

      lastEnd = end;
    }

    const trailing = line.slice(lastEnd).trim();
    if (trailing.length > 2) {
      results.push({
        rawName: trailing,
        displayName: extractDisplayName(trailing),
        faculty: null,
        hasHonorific: false,
      });
    }
  }

  return results;
}

/**
 * Parse time slot string into start/end Date objects.
 * Time format: "HH:MM - HH:MM"
 *
 * @param {string} timeStr
 * @param {Date|null} baseDate
 * @returns {{start:Date|null, end:Date|null}}
 */
function parseTimeSlot(timeStr, baseDate) {
  if (!timeStr || !baseDate) return { start: null, end: null };

  const match = timeStr.match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
  if (!match) return { start: null, end: null };

  const [, sh, sm, eh, em] = match.map((n) => parseInt(n, 10));

  const start = new Date(baseDate);
  start.setHours(sh, sm, 0, 0);

  const end = new Date(baseDate);
  end.setHours(eh, em, 0, 0);

  // Handle cases where end time appears earlier (e.g., 11:00 - 01:00 means next day)
  if (end < start) {
    end.setDate(end.getDate() + 1);
  }

  return { start, end };
}

/**
 * Determine assignment status relative to current time.
 *
 * @param {Date|null} start
 * @param {Date|null} end
 * @returns {'past'|'today'|'future'}
 */
function getAssignmentStatus(start, end) {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  if (!start || !end) return "future";

  if (end < now) return "past";

  const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  if (startDay.getTime() === today.getTime()) return "today";

  if (start > now) return "future";

  return "future";
}

/**
 * Determine overall day status based on assignments.
 *
 * @param {ScheduleAssignment[]} assignments
 * @returns {'past'|'today'|'future'}
 */
function getDayStatus(assignments) {
  const statuses = assignments.map((a) => a.status);
  if (statuses.every((s) => s === "past")) return "past";
  if (statuses.some((s) => s === "today")) return "today";
  return "future";
}

/**
 * Parse the metadata sheet to extract control rooms and notes.
 *
 * @param {any} workbook
 * @returns {{controlRooms:Map<string,string>, notes:string[]}}
 */
export function parseMetadataSheet(workbook) {
  const controlRooms = new Map();
  const notes = [];

  if (!workbook || !workbook.Sheets) return { controlRooms, notes };

  const sheetNames = workbook.SheetNames || [];
  const metaSheetName = sheetNames.find((n) => /الوان|كليات|الكليات|colors/i.test(n));
  if (!metaSheetName) return { controlRooms, notes };

  const sheet = workbook.Sheets[metaSheetName];
  if (!sheet) return { controlRooms, notes };

  const XLSX = window.XLSX;
  if (!XLSX || !XLSX.utils) return { controlRooms, notes };

  const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null;
  if (!range) return { controlRooms, notes };

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, range, defval: null });

  let headerRowIndex = -1;
  const abbrCols = [];
  const roomCols = [];

  for (let r = 0; r < Math.min(rows.length, 10); r++) {
    const row = rows[r];
    if (!row || !Array.isArray(row)) continue;
    for (let c = 0; c < row.length; c++) {
      const cell = cleanCell(row[c]);
      if (!cell) continue;
      if (/اختصارات|abbr|رمز|code|symbol/i.test(cell)) abbrCols.push({ row: r, col: c });
      if (/قاعة|room|hall|كنترول|control/i.test(cell)) roomCols.push({ row: r, col: c });
    }
    if (abbrCols.length > 0 || roomCols.length > 0) {
      headerRowIndex = r;
      break;
    }
  }

  let abbrCol = 7;
  let roomCol = 4;

  if (headerRowIndex >= 0) {
    const pick = (candidates) => {
      const byRow = candidates.filter((c) => c.row === headerRowIndex);
      return byRow.length > 0 ? byRow[0].col : candidates.length > 0 ? candidates[0].col : null;
    };
    const a = pick(abbrCols);
    const r = pick(roomCols);
    if (a != null) abbrCol = a;
    if (r != null) roomCol = r;
  }

  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !Array.isArray(row)) continue;
    const abbrCell = cleanCell(row[abbrCol]);
    const roomCell = cleanCell(row[roomCol]);
    if (abbrCell && roomCell) {
      controlRooms.set(abbrCell.toUpperCase(), roomCell);
    }
    for (let c = 0; c < row.length; c++) {
      const cell = cleanCell(row[c]);
      if (cell && cell.length > 20 && !controlRooms.has(cell.toUpperCase())) {
        notes.push(cell);
      }
    }
  }

  if (notes.length === 0) {
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !Array.isArray(row)) continue;
      for (let c = 0; c < row.length; c++) {
        const cell = cleanCell(row[c]);
        if (cell && cell.length > 20) notes.push(cell);
      }
    }
  }

  return { controlRooms, notes };
}

function findLastDataRow(rows) {
  for (let r = rows.length - 1; r >= 0; r--) {
    const row = rows[r];
    if (!row || !Array.isArray(row)) continue;
    if (row.some((cell) => cleanCell(cell) != null)) return r + 1;
  }
  return 0;
}

/**
 * Parse a single schedule sheet.
 *
 * @param {any[][]} rows
 * @param {string} sheetName
 * @param {Map<string,string>} controlRooms
 * @returns {DaySchedule|null}
 */
export function parseScheduleSheet(rows, sheetName, controlRooms) {
  const parsedName = parseSheetName(sheetName);
  if (!parsedName) return null;

  const date = buildDateFromParsed(parsedName);
  const dateDisplay = formatDateDisplay(date);

  const assignments = [];
  const lastRow = findLastDataRow(rows);

  for (let r = 0; r < lastRow; r++) {
    const row = rows[r];
    if (!row || !Array.isArray(row)) continue;

    const room = cleanCell(row[0]);
    const time = cleanCell(row[1]);
    const faculty = cleanCell(row[2]);

    const hasContext = room || time || faculty;
    if (!hasContext) continue;

    /** @type {SubHall[]} */
    const subHalls = [];

    for (let c = 3; c < row.length; c++) {
      const cell = cleanCell(row[c]);
      if (!cell) continue;

      const proctors = splitProctorCell(cell);
      if (proctors.length === 0) continue;

      subHalls.push({
        groupIndex: c - 2, // 1-based sub-hall number
        proctors,
      });
    }

    if (subHalls.length === 0) continue;

    const isIncomplete = !room || !time || !faculty;

    let controlRoom = UNSPECIFIED;
    if (faculty) {
      const lookup = controlRooms.get(faculty.toUpperCase());
      if (lookup) controlRoom = lookup;
    }

    let inferredFaculty = faculty;
    if (!inferredFaculty && subHalls.length > 0) {
      const firstProctor = subHalls[0].proctors[0];
      if (firstProctor && firstProctor.faculty) {
        inferredFaculty = firstProctor.faculty;
        const lookup = controlRooms.get(firstProctor.faculty.toUpperCase());
        if (lookup) controlRoom = lookup;
      }
    }

    const { start, end } = parseTimeSlot(time || "", date);
    const status = getAssignmentStatus(start, end);

    assignments.push({
      room: room || UNSPECIFIED,
      time: time || UNSPECIFIED,
      faculty: inferredFaculty || UNSPECIFIED,
      controlRoom,
      subHalls,
      rawRow: r + 1,
      incomplete: isIncomplete,
      startTime: start,
      endTime: end,
      status,
    });
  }

  if (assignments.length === 0) return null;

  const dayStatus = getDayStatus(assignments);

  return {
    dayArabic: parsedName.day,
    dayEnglish: DAY_MAP[parsedName.day] || parsedName.day,
    periodArabic: parsedName.period,
    periodEnglish: PERIOD_MAP[parsedName.period] || parsedName.period,
    date,
    dateDisplay,
    originalName: sheetName,
    assignments,
    dayStatus,
  };
}

/**
 * Parse the entire workbook.
 *
 * @param {any} workbook
 * @returns {ParsedProctoringData}
 */
export function parseProctoringWorkbook(workbook) {
  const days = [];
  const allNamesSet = new Set();
  const allNamesOriginal = new Map();
  let totalAssignments = 0;

  if (!workbook || !workbook.SheetNames || !workbook.Sheets) {
    return { days: [], controlRooms: new Map(), allProctorNames: [], notes: [], totalAssignments: 0 };
  }

  const { controlRooms, notes } = parseMetadataSheet(workbook);

  const XLSX = window.XLSX;
  if (!XLSX || !XLSX.utils) {
    return { days: [], controlRooms: new Map(), allProctorNames: [], notes: [], totalAssignments: 0 };
  }

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null;
    if (!range) continue;
    if (/الوان|كليات|الكليات|colors/i.test(sheetName)) continue;

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, range, defval: null });
    const daySchedule = parseScheduleSheet(rows, sheetName, controlRooms);
    if (!daySchedule) continue;

    days.push(daySchedule);
    totalAssignments += daySchedule.assignments.length;

    for (const assignment of daySchedule.assignments) {
      for (const subHall of assignment.subHalls) {
        for (const p of subHall.proctors) {
          const norm = normalizeProctorName(p.rawName);
          if (!norm) continue;
          if (!allNamesSet.has(norm)) {
            allNamesSet.add(norm);
            allNamesOriginal.set(norm, p.displayName);
          }
        }
      }
    }
  }

  const periodOrder = { اولي: 1, تانيه: 2, تالته: 3, رابعه: 4 };
  days.sort((a, b) => {
    if (a.date && b.date) return a.date.getTime() - b.date.getTime();
    if (a.date && !b.date) return -1;
    if (!a.date && b.date) return 1;
    return a.originalName.localeCompare(b.originalName);
  });
  days.sort((a, b) => {
    if (a.date && b.date && a.date.getTime() === b.date.getTime()) {
      return (periodOrder[a.periodArabic] || 99) - (periodOrder[b.periodArabic] || 99);
    }
    return 0;
  });

  const allProctorNames = Array.from(allNamesOriginal.values());
  allProctorNames.sort((a, b) => a.localeCompare(b, "ar"));

  return { days, controlRooms, allProctorNames, notes, totalAssignments };
}

/**
 * Filter schedule data for a specific proctor name.
 * Uses fuzzy matching to handle spelling variations (e.g., "حمد حاتم" vs "محمد حاتم").
 *
 * @param {ParsedProctoringData} data
 * @param {string} query
 * @returns {{day:DaySchedule, assignments:ScheduleAssignment[]}[]}
 */
export function filterScheduleForProctor(data, query) {
  if (!data || !query) return [];
  const normalizedQuery = normalizeProctorName(query);
  if (!normalizedQuery) return [];

  // Collect all name variants (exact/substring/typo) for this query.
  const similarNames = findSimilarNames(data.allProctorNames, query);
  const normalizedSimilarNames = new Set(similarNames.map(n => normalizeProctorName(n)).filter(Boolean));

  const results = [];
  for (const day of data.days) {
    const matched = [];
    for (const assignment of day.assignments) {
      const hasMatch = assignment.subHalls.some((sh) =>
        sh.proctors.some((p) => {
          const norm = normalizeProctorName(p.rawName);
          if (!norm) return false;

          if (norm === normalizedQuery || norm.includes(normalizedQuery)) return true;

          for (const similarNorm of normalizedSimilarNames) {
            if (norm === similarNorm || areNamesSimilar(p.rawName, query)) return true;
          }

          return false;
        })
      );
      if (hasMatch) matched.push(assignment);
    }
    if (matched.length > 0) results.push({ day, assignments: matched });
  }
  return results;
}

/**
 * Get autocomplete suggestions for a query.
 * Returns exact/substring matches first, then word-level typo variants.
 * Different people who share a first name are shown as separate entries.
 *
 * @param {ParsedProctoringData} data
 * @param {string} query
 * @param {number} [limit=10]
 * @returns {string[]}
 */
export function getNameSuggestions(data, query, limit = 10) {
  if (!data || !query) return [];
  const normalizedQuery = normalizeProctorName(query);
  if (!normalizedQuery) return data.allProctorNames.slice(0, limit);

  // Exact / substring matches.
  const exactMatches = data.allProctorNames.filter((name) => {
    const norm = normalizeProctorName(name);
    return norm && (norm.includes(normalizedQuery) || normalizedQuery.includes(norm));
  });

  // Word-level typo variants not already covered by substring matching.
  const fuzzyMatches = data.allProctorNames.filter((name) => {
    if (exactMatches.includes(name)) return false;
    return areNamesSimilar(name, query);
  });

  // Deduplicate typo variants: if two names are word-level typos of each other,
  // keep only the longer (more complete) spelling.
  const combined = [...exactMatches, ...fuzzyMatches];
  const uniqueNames = [];
  const seen = new Set();

  for (const name of combined) {
    const norm = normalizeProctorName(name);
    if (!norm || seen.has(norm)) continue;

    let mergedIntoExisting = false;
    for (let i = 0; i < uniqueNames.length; i++) {
      if (areNamesSimilar(name, uniqueNames[i])) {
        if (name.length > uniqueNames[i].length) uniqueNames[i] = name;
        mergedIntoExisting = true;
        break;
      }
    }

    if (!mergedIntoExisting) {
      uniqueNames.push(name);
      seen.add(norm);
    }
  }

  return uniqueNames.slice(0, limit);
}
