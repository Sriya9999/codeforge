/*
 * Grading Console — core logic.
 *
 * Pure functions only: no DOM, no globals besides the exported namespace.
 * Loaded as a classic <script> in the browser (window.GradingCore) and via
 * require() in Node tests. Every rule that decides a student's grade lives
 * here so it can be tested in isolation from the interface.
 */
(function (root) {
  "use strict";

  const GRADES = ["A", "A-", "B", "B-", "C", "C-", "D", "E"];
  const DEFAULT_BANDS = [
    { grade: "A", min: 80, max: 100 },
    { grade: "A-", min: 70, max: 79 },
    { grade: "B", min: 60, max: 69 },
    { grade: "B-", min: 50, max: 59 },
    { grade: "C", min: 40, max: 49 },
    { grade: "C-", min: 30, max: 39 },
    { grade: "D", min: 20, max: 29 },
    { grade: "E", min: 0, max: 19 },
  ];
  const MARK_MIN = 0;
  const MARK_MAX = 100;

  /* ------------------------------------------------------------------ *
   * Column detection
   * ------------------------------------------------------------------ */

  // Lower-cases, folds curly quotes, and strips everything but letters/digits,
  // so "Student’s BITS ID", "Student's BITS ID" and " bits id " compare equal
  // on their essential characters.
  function normaliseHeader(h) {
    return String(h == null ? "" : h)
      .toLowerCase()
      .replace(/[‘’ʼ`´]/g, "'")
      .replace(/[^a-z0-9]/g, "");
  }

  const COLUMNS = {
    id: { label: "Student’s BITS ID", test: (n) => n.includes("bitsid") },
    course: { label: "Course", test: (n) => n === "course" || n === "coursename" || n === "coursetitle" },
    marks: { label: "Total Marks", test: (n) => n === "totalmarks" || n === "total" || n === "totalmark" },
  };

  function detectColumns(headerRow) {
    const found = {};
    const extras = [];
    (headerRow || []).forEach((h, i) => {
      const n = normaliseHeader(h);
      if (!n) return;
      const key = Object.keys(COLUMNS).find((k) => COLUMNS[k].test(n));
      if (key && found[key] === undefined) found[key] = i;
      else extras.push(String(h).trim());
    });
    const missing = Object.keys(COLUMNS).filter((k) => found[k] === undefined).map((k) => COLUMNS[k].label);
    return { index: found, missing, extras };
  }

  /* ------------------------------------------------------------------ *
   * Row parsing & validation
   * ------------------------------------------------------------------ */

  const isBlank = (v) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  const cleanText = (v) => (isBlank(v) ? "" : String(v).replace(/\s+/g, " ").trim());

  // Returns { value, kind } where kind is one of:
  //   ok | text-number | decimal | blank | non-numeric | out-of-range
  function parseMark(raw) {
    if (isBlank(raw)) return { value: null, kind: "blank" };
    let n;
    let fromText = false;
    if (typeof raw === "number") n = raw;
    else {
      const s = String(raw).trim();
      if (!/^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(s)) return { value: null, kind: "non-numeric" };
      n = Number(s);
      fromText = true;
    }
    if (!Number.isFinite(n)) return { value: null, kind: "non-numeric" };
    if (n < MARK_MIN || n > MARK_MAX) return { value: n, kind: "out-of-range" };
    if (!Number.isInteger(n)) return { value: n, kind: "decimal" };
    return { value: n, kind: fromText ? "text-number" : "ok" };
  }

  const ISSUE_TEXT = {
    "missing-id": "BITS ID is blank",
    "missing-course": "Course is blank",
    blank: "Total Marks is blank",
    "non-numeric": "Total Marks is not a number",
    "out-of-range": "Total Marks is outside 0–100",
    decimal: "Total Marks is not a whole number",
    "duplicate-conflict": "Same BITS ID appears again in this course with a different mark",
    "duplicate-exact": "Exact duplicate of an earlier row — counted once",
  };

  /**
   * Parse a sheet given as an array of rows (arrays of cell values), with the
   * header in the first non-empty row.
   *
   * Returns:
   *   { ok:false, error }                         when the file can't be graded at all
   *   { ok:true, records, courses, columns, … }   otherwise
   *
   * Each record: { row, id, course, raw, mark, status, issues[] }
   *   status: "valid"     → will be graded
   *           "attention" → excluded until resolved (never silently graded)
   *           "duplicate" → exact repeat of a valid row; ignored
   */
  function parseSheet(rows, opts) {
    const options = Object.assign({ roundDecimals: false }, opts);
    const firstIdx = (rows || []).findIndex((r) => Array.isArray(r) && r.some((c) => !isBlank(c)));
    if (firstIdx === -1) return { ok: false, error: { code: "empty", message: "The sheet is empty — no header row and no student records were found." } };

    const columns = detectColumns(rows[firstIdx]);
    if (columns.missing.length) {
      const seen = rows[firstIdx].filter((c) => !isBlank(c)).map((c) => String(c).trim());
      return {
        ok: false,
        error: {
          code: "columns",
          message: `Missing required column${columns.missing.length > 1 ? "s" : ""}: ${columns.missing.join(", ")}.`,
          found: seen,
          missing: columns.missing,
        },
      };
    }

    const { id: ci, course: cc, marks: cm } = columns.index;
    const records = [];
    const notes = { textNumbers: 0, whitespaceFixed: 0, rounded: [] };

    for (let r = firstIdx + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const rawId = row[ci], rawCourse = row[cc], rawMark = row[cm];
      if (isBlank(rawId) && isBlank(rawCourse) && isBlank(rawMark)) continue; // blank spacer row

      const id = cleanText(rawId);
      const course = cleanText(rawCourse);
      if (!isBlank(rawCourse) && String(rawCourse) !== course) notes.whitespaceFixed++;

      let pm = parseMark(rawMark);
      if (pm.kind === "decimal" && options.roundDecimals) {
        const rounded = Math.floor(pm.value + 0.5); // half-up; marks are non-negative here
        notes.rounded.push({ row: r + 1, id, course, from: pm.value, to: rounded });
        pm = { value: rounded, kind: "ok" };
      }
      if (pm.kind === "text-number") notes.textNumbers++;

      const issues = [];
      if (!id) issues.push("missing-id");
      if (!course) issues.push("missing-course");
      if (!["ok", "text-number"].includes(pm.kind)) issues.push(pm.kind);

      records.push({
        row: r + 1, // 1-based spreadsheet row, as the instructor sees it in Excel
        id,
        course,
        raw: rawMark,
        mark: issues.length ? null : pm.value,
        status: issues.length ? "attention" : "valid",
        issues,
      });
    }

    // Duplicate BITS IDs within a course. Exact repeats are harmless and are
    // counted once; conflicting marks can't be resolved automatically, so
    // every row for that ID is held back for the instructor.
    const byKey = new Map();
    records.forEach((rec) => {
      if (!rec.id || !rec.course) return;
      const key = rec.course + "\u0000" + rec.id.toUpperCase();
      if (!byKey.has(key)) byKey.set(key, []);
      byKey.get(key).push(rec);
    });
    byKey.forEach((group) => {
      if (group.length < 2) return;
      const marks = new Set(group.map((g) => (g.status === "valid" ? g.mark : "∅" + String(g.raw))));
      if (marks.size === 1 && group[0].status === "valid") {
        group.slice(1).forEach((g) => { g.status = "duplicate"; g.issues.push("duplicate-exact"); });
      } else {
        group.forEach((g) => {
          g.status = "attention";
          g.mark = null;
          if (!g.issues.includes("duplicate-conflict")) g.issues.push("duplicate-conflict");
        });
      }
    });

    const courses = summariseCourses(records);
    const counts = {
      rows: records.length,
      valid: records.filter((r) => r.status === "valid").length,
      attention: records.filter((r) => r.status === "attention").length,
      duplicates: records.filter((r) => r.status === "duplicate").length,
      decimals: records.filter((r) => r.issues.includes("decimal")).length,
      noCourse: records.filter((r) => !r.course).length,
    };
    return { ok: true, records, courses, columns, counts, notes };
  }

  function summariseCourses(records) {
    const map = new Map();
    records.forEach((r) => {
      if (!r.course) return;
      if (!map.has(r.course)) map.set(r.course, { name: r.course, valid: 0, attention: 0, duplicates: 0, total: 0 });
      const c = map.get(r.course);
      c.total++;
      if (r.status === "valid") c.valid++;
      else if (r.status === "attention") c.attention++;
      else c.duplicates++;
    });
    return [...map.values()].sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  }

  /* ------------------------------------------------------------------ *
   * Statistics
   * ------------------------------------------------------------------ */

  function stats(marks) {
    const m = marks.filter((x) => typeof x === "number" && Number.isFinite(x)).slice().sort((a, b) => a - b);
    const n = m.length;
    if (!n) return { n: 0, min: null, max: null, mean: null, median: null, sd: null };
    const sum = m.reduce((a, b) => a + b, 0);
    const mean = sum / n;
    const median = n % 2 ? m[(n - 1) / 2] : (m[n / 2 - 1] + m[n / 2]) / 2;
    const sd = Math.sqrt(m.reduce((a, b) => a + (b - mean) ** 2, 0) / n);
    return { n, min: m[0], max: m[n - 1], mean, median, sd };
  }

  // Ten bins: 0–9, 10–19, …, 80–89, 90–100 (the last bin includes 100).
  function histogram(marks, width) {
    const w = width || 10;
    const count = Math.ceil((MARK_MAX + 1) / w) - (MARK_MAX % w === 0 ? 1 : 0);
    const bins = Array.from({ length: count }, (_, i) => ({ lo: i * w, hi: i === count - 1 ? MARK_MAX : i * w + w - 1, n: 0 }));
    marks.forEach((x) => { bins[Math.min(count - 1, Math.floor(x / w))].n++; });
    return bins;
  }

  /* ------------------------------------------------------------------ *
   * Grade bands
   * ------------------------------------------------------------------ */

  const cloneBands = (b) => b.map((x) => ({ grade: x.grade, min: x.min, max: x.max }));

  /**
   * Validate a grading scheme. Returns a list of problems, each with a
   * precise, human sentence. An empty list means: every whole mark 0–100
   * maps to exactly one grade, and grades are in descending order.
   */
  function validateBands(bands) {
    const problems = [];
    const add = (type, message, grades) => problems.push({ type, message, grades: grades || [] });

    bands.forEach((b) => {
      ["min", "max"].forEach((k) => {
        const v = b[k];
        const label = k === "min" ? "lower bound" : "upper bound";
        if (v === null || v === undefined || v === "" || !Number.isFinite(Number(v))) add("invalid", `${b.grade}: ${label} is empty.`, [b.grade]);
        else if (!Number.isInteger(Number(v))) add("invalid", `${b.grade}: ${label} ${v} must be a whole number.`, [b.grade]);
        else if (v < MARK_MIN || v > MARK_MAX) add("range", `${b.grade}: ${label} ${v} is outside 0–100.`, [b.grade]);
      });
    });
    if (problems.length) return problems; // coverage maths below needs clean integers

    bands.forEach((b) => {
      if (b.min > b.max) add("reversed", `${b.grade}: lower bound ${b.min} is above upper bound ${b.max}.`, [b.grade]);
    });

    for (let i = 0; i < bands.length - 1; i++) {
      const hi = bands[i], lo = bands[i + 1];
      if (hi.min > hi.max || lo.min > lo.max) continue;
      if (lo.max >= hi.min && lo.min <= hi.max) {
        const a = Math.max(lo.min, hi.min), z = Math.min(lo.max, hi.max);
        add("overlap", `${hi.grade} and ${lo.grade} overlap on ${span(a, z)}.`, [hi.grade, lo.grade]);
      } else if (lo.min > hi.max) {
        add("order", `${lo.grade} (${lo.min}–${lo.max}) sits above ${hi.grade} (${hi.min}–${hi.max}); grades must descend from A to E.`, [hi.grade, lo.grade]);
      }
    }

    // Coverage: walk every whole mark and report contiguous gaps / overlaps
    // that the neighbour checks above did not already describe.
    const owners = Array.from({ length: MARK_MAX + 1 }, () => []);
    bands.forEach((b) => { for (let x = Math.max(MARK_MIN, b.min); x <= Math.min(MARK_MAX, b.max); x++) owners[x].push(b.grade); });
    let start = null;
    for (let x = 0; x <= MARK_MAX + 1; x++) {
      const gap = x <= MARK_MAX && owners[x].length === 0;
      if (gap && start === null) start = x;
      if (!gap && start !== null) {
        const above = bands.filter((b) => b.min > x - 1).map((b) => b.grade);
        add("gap", `No grade covers ${span(start, x - 1)}${gapHint(bands, start, x - 1)}.`, above.slice(-1));
        start = null;
      }
    }
    const described = new Set(problems.filter((p) => p.type === "overlap").map((p) => p.grades.join("|")));
    for (let x = 0; x <= MARK_MAX; x++) {
      if (owners[x].length > 1) {
        const key = owners[x].join("|");
        if (!described.has(key)) {
          let z = x;
          while (z + 1 <= MARK_MAX && owners[z + 1].join("|") === key) z++;
          add("overlap", `${owners[x].join(" and ")} overlap on ${span(x, z)}.`, owners[x]);
          described.add(key);
          x = z;
        }
      }
    }
    return problems;
  }

  function span(a, b) { return a === b ? `mark ${a}` : `marks ${a}–${b}`; }
  function gapHint(bands, a, b) {
    const above = bands.find((x) => x.min === b + 1);
    const below = bands.find((x) => x.max === a - 1);
    if (above && below) return ` (between ${below.grade} and ${above.grade})`;
    if (a === MARK_MIN) return " (bottom of the scale)";
    if (b === MARK_MAX) return " (top of the scale)";
    return "";
  }

  function gradeFor(mark, bands) {
    for (const b of bands) if (mark >= b.min && mark <= b.max) return b.grade;
    return null;
  }

  function gradeDistribution(marks, bands) {
    const out = bands.map((b) => ({ grade: b.grade, min: b.min, max: b.max, n: 0 }));
    let unassigned = 0;
    marks.forEach((m) => {
      const i = bands.findIndex((b) => m >= b.min && m <= b.max);
      if (i === -1) unassigned++;
      else out[i].n++;
    });
    return { bands: out, unassigned };
  }

  // Students within `within` marks below the lower bound of the next grade up.
  function nearBoundary(records, bands, within) {
    const w = within == null ? 1 : within;
    const res = [];
    records.forEach((r) => {
      const i = bands.findIndex((b) => r.mark >= b.min && r.mark <= b.max);
      if (i <= 0) return;
      const up = bands[i - 1];
      const gap = up.min - r.mark;
      if (gap > 0 && gap <= w) res.push({ id: r.id, mark: r.mark, grade: bands[i].grade, next: up.grade, cutoff: up.min, gap });
    });
    return res.sort((a, b) => a.gap - b.gap || b.mark - a.mark);
  }

  /* ------------------------------------------------------------------ *
   * Insights — factual sentences derived only from the data
   * ------------------------------------------------------------------ */

  function insights(marks, bands) {
    const out = [];
    const s = stats(marks);
    if (!s.n) return out;
    if (s.n < 5) {
      out.push(`Only ${s.n} student${s.n > 1 ? "s" : ""} in this course — distribution statistics carry little meaning.`);
    }

    if (s.sd === 0) {
      out.push(`Every student scored ${s.min}.`);
    } else {
      // Densest 20-mark window over the ten decade bins.
      const bins = histogram(marks);
      let best = 0;
      for (let i = 1; i < bins.length - 1; i++) if (bins[i].n + bins[i + 1].n > bins[best].n + bins[best + 1].n) best = i;
      const inWin = bins[best].n + bins[best + 1].n;
      const pct = Math.round((inWin / s.n) * 100);
      out.push(`${pct}% of students (${inWin} of ${s.n}) scored between ${bins[best].lo} and ${bins[best + 1].hi}.`);

      const diff = s.mean - s.median;
      if (s.n >= 20 && Math.abs(diff) >= 1) {
        out.push(`The mean (${fmt(s.mean)}) is ${fmt(Math.abs(diff))} ${diff < 0 ? "below" : "above"} the median (${fmt(s.median)}): a ${diff < 0 ? "tail of lower" : "tail of higher"} marks pulls the average ${diff < 0 ? "down" : "up"}.`);
      }
    }

    if (!validateBands(bands).length) {
      const d = gradeDistribution(marks, bands).bands;
      const top = d.reduce((a, b) => (b.n > a.n ? b : a));
      out.push(`${top.grade} is the most common grade: ${top.n} student${top.n === 1 ? "" : "s"} (${Math.round((top.n / s.n) * 100)}%).`);
      const empty = d.filter((b) => b.n === 0).map((b) => b.grade);
      if (empty.length && empty.length < d.length) out.push(`No student currently receives ${listJoin(empty)}.`);
    }
    return out;
  }

  function fmt(x, dp) {
    if (x === null || x === undefined || !Number.isFinite(x)) return "—";
    const d = dp == null ? 1 : dp;
    return Number.isInteger(x) ? String(x) : x.toFixed(d);
  }
  function listJoin(a) { return a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " or " + a[a.length - 1]; }

  /* ------------------------------------------------------------------ *
   * Export
   * ------------------------------------------------------------------ */

  // RFC 4180 quoting plus neutralising spreadsheet formula prefixes so a cell
  // like "=HYPERLINK(...)" is shown as text, never executed, when the CSV is
  // opened in Excel / Sheets.
  function csvCell(v) {
    let s = v === null || v === undefined ? "" : String(v);
    if (typeof v !== "number" && /^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  /**
   * Build the export CSV. Layout matches the original console so downstream
   * consumers keep working:
   *
   *   Instructor,<name>
   *   Course,<course>
   *   <blank>
   *   BITS ID,Total Marks,Grade
   *   …one row per graded student…
   *
   * Throws if any student would be left without a grade — the export is
   * never allowed to silently drop a student.
   */
  function buildCsv({ instructor, course, records, bands }) {
    const problems = validateBands(bands);
    if (problems.length) throw new Error("Grade ranges are invalid: " + problems[0].message);
    const lines = [
      ["Instructor", instructor].map(csvCell).join(","),
      ["Course", course].map(csvCell).join(","),
      "",
      "BITS ID,Total Marks,Grade",
    ];
    records.forEach((r) => {
      const g = gradeFor(r.mark, bands);
      if (!g) throw new Error(`Student ${r.id} (mark ${r.mark}) has no grade.`);
      lines.push([r.id, r.mark, g].map(csvCell).join(","));
    });
    return lines.join("\r\n") + "\r\n";
  }

  function exportFilename(course, date) {
    const d = date || new Date();
    const slug = String(course).normalize("NFKD").replace(/[^\w\s-]/g, "").trim().replace(/[\s_]+/g, "-").toLowerCase().slice(0, 60) || "course";
    const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    return `grades_${slug}_${stamp}.csv`;
  }

  const api = {
    GRADES, DEFAULT_BANDS, ISSUE_TEXT,
    normaliseHeader, detectColumns, parseMark, parseSheet,
    stats, histogram,
    cloneBands, validateBands, gradeFor, gradeDistribution, nearBoundary,
    insights, fmt, csvCell, buildCsv, exportFilename,
  };
  root.GradingCore = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
