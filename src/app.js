/*
 * Grading Console — interface.
 *
 * All grading rules live in core.js (GradingCore). This file owns state,
 * rendering and interaction. DOM is built with createElement / textContent,
 * never innerHTML with data, so names from a spreadsheet can't inject markup.
 */
(function () {
  "use strict";
  const C = window.GradingCore;
  const MAX_FILE_BYTES = 25 * 1024 * 1024;
  const ACCEPTED = [".xlsx", ".xls"];
  const STEPS = [
    { key: "import", n: "01", label: "Import" },
    { key: "review", n: "02", label: "Review" },
    { key: "grade", n: "03", label: "Grade" },
    { key: "export", n: "04", label: "Export" },
  ];

  // localStorage can throw (private mode, blocked storage); it only ever
  // remembers the instructor's name, so failures are ignored.
  const storage = {
    get(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { window.localStorage.setItem(k, v); } catch (e) { /* ignore */ } },
  };

  /* ------------------------------------------------------------------ *
   * State
   * ------------------------------------------------------------------ */
  const state = {
    step: "import",
    load: null,           // { phase: "reading"|"parsing", name, pct } while a file is in flight
    loadToken: 0,         // guards against an older file finishing after a newer one
    file: null,           // { name, size, sheet, sheetCount }
    rows: null,           // raw rows, kept so decimals can be re-parsed on request
    parsed: null,         // result of C.parseSheet (ok or error)
    fileError: null,      // { title, detail, found? }
    roundDecimals: false,
    course: null,
    bandsByCourse: new Map(),  // kept across re-imports: fixing the file never loses cutoffs
    linkBounds: true,
    exports: new Map(),        // course → { at, signature, count, filename }
    importedAt: null,
    instructor: storage.get("gc.instructor") || "",
    ack: new Set(),            // courses where the instructor acknowledged excluded records
    reviewFilter: "",
  };

  /* ------------------------------------------------------------------ *
   * Small DOM helpers
   * ------------------------------------------------------------------ */
  function h(tag, props, ...kids) {
    const el = document.createElement(tag);
    applyProps(el, props);
    append(el, kids);
    return el;
  }
  const SVGNS = "http://www.w3.org/2000/svg";
  function s(tag, attrs, ...kids) {
    const el = document.createElementNS(SVGNS, tag);
    for (const k in attrs || {}) if (attrs[k] !== undefined && attrs[k] !== null) el.setAttribute(k, attrs[k]);
    append(el, kids);
    return el;
  }
  function applyProps(el, props) {
    for (const k in props || {}) {
      const v = props[k];
      if (v === undefined || v === null || v === false) continue;
      if (k === "class") el.className = v;
      else if (k === "text") el.textContent = v;
      else if (k === "style") el.setAttribute("style", v);
      else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
      else if (k in el && typeof v !== "string") el[k] = v;
      else el.setAttribute(k, v === true ? "" : v);
    }
  }
  function append(el, kids) {
    kids.flat(Infinity).forEach((c) => {
      if (c === null || c === undefined || c === false) return;
      el.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
    });
  }
  const $ = (sel, root) => (root || document).querySelector(sel);
  const plural = (n, one, many) => `${n.toLocaleString()} ${n === 1 ? one : many || one + "s"}`;
  const fmt = (x, dp) => C.fmt(x, dp);
  const fmtBytes = (b) => (b < 1024 ? b + " B" : b < 1048576 ? (b / 1024).toFixed(0) + " KB" : (b / 1048576).toFixed(1) + " MB");
  const timeOf = (d) => d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });



  /* ------------------------------------------------------------------ *
   * Derived data for the selected course
   * ------------------------------------------------------------------ */
  function courseRecords(course) {
    if (!state.parsed || !state.parsed.ok) return [];
    return state.parsed.records.filter((r) => r.course === course);
  }
  function courseView(course) {
    const all = courseRecords(course);
    const valid = all.filter((r) => r.status === "valid");
    const attention = all.filter((r) => r.status === "attention");
    const duplicates = all.filter((r) => r.status === "duplicate");
    const marks = valid.map((r) => r.mark);
    return { name: course, all, valid, attention, duplicates, marks, stats: C.stats(marks) };
  }
  function bandsFor(course) {
    if (!state.bandsByCourse.has(course)) state.bandsByCourse.set(course, C.cloneBands(C.DEFAULT_BANDS));
    return state.bandsByCourse.get(course);
  }
  function signature(course) {
    const v = courseView(course);
    const str = JSON.stringify(bandsFor(course)) + "|" + v.valid.map((r) => r.id + ":" + r.mark).join(",");
    let x = 5381;
    for (let i = 0; i < str.length; i++) x = ((x << 5) + x + str.charCodeAt(i)) | 0;
    return x;
  }
  function exportState(course) {
    const e = state.exports.get(course);
    if (!e) return null;
    return { ...e, stale: e.signature !== signature(course) };
  }
  const hasData = () => !!(state.parsed && state.parsed.ok && state.parsed.records.length);
  const canGrade = () => hasData() && state.course && courseView(state.course).valid.length > 0;

  /* ------------------------------------------------------------------ *
   * File import
   * ------------------------------------------------------------------ */
  function handleFile(file) {
    if (!file) return;
    const token = ++state.loadToken;
    state.fileError = null;
    const ext = (file.name.match(/\.[^.]+$/) || [""])[0].toLowerCase();
    if (!ACCEPTED.includes(ext)) {
      return failImport(token, "This isn’t an Excel workbook",
        `“${file.name}” is ${ext ? "a " + ext + " file" : "not an Excel file"}. Save the marks from Excel as an .xlsx workbook and try again.`);
    }
    if (file.size === 0) return failImport(token, "The file is empty", `“${file.name}” has 0 bytes. It may not have finished saving or downloading.`);
    if (file.size > MAX_FILE_BYTES) return failImport(token, "The file is too large", `“${file.name}” is ${fmtBytes(file.size)}. Files up to ${fmtBytes(MAX_FILE_BYTES)} are supported.`);

    state.load = { phase: "reading", name: file.name, pct: 0 };
    render();
    const reader = new FileReader();
    reader.onprogress = (e) => {
      if (token !== state.loadToken || !e.lengthComputable) return;
      state.load.pct = Math.round((e.loaded / e.total) * 100);
      const bar = $("#loadBar");
      if (bar) bar.style.width = state.load.pct + "%";
    };
    reader.onerror = () => failImport(token, "The file couldn’t be read", "Your browser was unable to read the file. Check that it isn’t open and locked by another program, then try again.");
    reader.onload = () => {
      if (token !== state.loadToken) return;
      state.load = { phase: "parsing", name: file.name, pct: 100 };
      render();
      // Yield so the "Checking rows" state paints before a large parse blocks the thread.
      setTimeout(() => parseWorkbook(token, file, reader.result), 30);
    };
    reader.readAsArrayBuffer(file);
  }

  function parseWorkbook(token, file, buffer) {
    if (token !== state.loadToken) return;
    let wb;
    try {
      wb = XLSX.read(buffer, { type: "array", dense: true, cellHTML: false, cellFormula: false });
    } catch (err) {
      console.warn("Workbook parse failed:", err);
      // Encrypted OOXML files are CFB containers that SheetJS reports as "password-protected".
      // (A damaged ZIP can report "Unsupported ZIP encryption" — that is corruption, not a password.)
      const locked = /password/i.test(String(err && err.message));
      return failImport(token, locked ? "The workbook is password-protected" : "This file couldn’t be opened as a workbook",
        locked ? "Remove the password in Excel (File → Info → Protect Workbook), save, and import again."
               : `“${file.name}” may be damaged, or it isn’t really an Excel file despite its name. Open it in Excel, choose Save As → Excel Workbook (.xlsx), and import the new copy.`);
    }

    // Use the first sheet that has the three required columns.
    let chosen = null, firstResult = null;
    for (const name of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
      const res = C.parseSheet(rows, { roundDecimals: false });
      if (!firstResult && !(res.error && res.error.code === "empty")) firstResult = { name, rows, res };
      if (res.ok) { chosen = { name, rows, res }; break; }
    }
    if (!chosen) {
      const r = firstResult && firstResult.res;
      if (!r) return failImport(token, "The workbook is empty", `No sheet in “${file.name}” contains a header row or any student records.`);
      return failImport(token, "The required columns weren’t found",
        r.error.message + ` The first row of sheet “${firstResult.name}” should name the columns exactly as shown under Expected format.`,
        r.error.found);
    }

    state.file = { name: file.name, size: file.size, sheet: chosen.name, sheetCount: wb.SheetNames.length };
    state.rows = chosen.rows;
    state.roundDecimals = false;
    state.load = null;
    state.importedAt = new Date();
    state.ack.clear();
    reparse();
    const courses = state.parsed.courses;
    const keep = state.course && courses.some((c) => c.name === state.course) ? state.course : null;
    state.course = keep || (courses.find((c) => c.valid > 0) || courses[0] || {}).name || null;
    render();
    const kept = courses.filter((c) => state.bandsByCourse.has(c.name) && JSON.stringify(state.bandsByCourse.get(c.name)) !== JSON.stringify(C.DEFAULT_BANDS));
    if (kept.length) toast(`Imported. Your grade ranges for ${kept.length === 1 ? kept[0].name : plural(kept.length, "course")} were kept.`);
  }

  function reparse() {
    state.parsed = C.parseSheet(state.rows, { roundDecimals: state.roundDecimals });
  }

  function failImport(token, title, detail, found) {
    if (token !== state.loadToken) return;
    state.load = null;
    state.fileError = { title, detail, found };
    state.step = "import";
    render();
  }

  function loadSample() {
    fetch("fixtures/sample-marks.xlsx")
      .then((r) => { if (!r.ok) throw new Error(r.status); return r.blob(); })
      .then((b) => handleFile(new File([b], "sample-marks.xlsx", { type: b.type })))
      .catch(() => toast("The sample file couldn’t be loaded."));
  }

  function downloadTemplate() {
    const ws = XLSX.utils.aoa_to_sheet([
      ["Student’s BITS ID", "Course", "Total Marks"],
      ["2024A7PS0001P", "Course A", 82],
      ["2024A7PS0002P", "Course A", 71],
      ["2024A7PS0003P", "Course A", 64],
    ]);
    ws["!cols"] = [{ wch: 20 }, { wch: 28 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Marks");
    XLSX.writeFile(wb, "marks-template.xlsx");
  }

  /* ------------------------------------------------------------------ *
   * Navigation
   * ------------------------------------------------------------------ */
  function stepAvailable(key) {
    if (key === "import") return true;
    if (key === "review") return hasData() || !!(state.parsed && state.parsed.ok);
    return canGrade();
  }
  function go(key) {
    if (!stepAvailable(key)) return;
    state.step = key;
    render();
    const main = $("#main");
    main.focus({ preventScroll: true });
    window.scrollTo({ top: 0 });
  }
  function selectCourse(name) {
    state.course = name;
    if (!canGrade() && (state.step === "grade" || state.step === "export")) state.step = "review";
    render();
  }

  /* ------------------------------------------------------------------ *
   * Render root
   * ------------------------------------------------------------------ */
  function render() {
    renderSteps();
    renderFileChip();
    renderRail();
    const main = $("#main");
    main.replaceChildren();
    const view = { import: viewImport, review: viewReview, grade: viewGrade, export: viewExport }[state.step];
    main.appendChild(view());
    if (state.step === "grade") updateGrade();
    document.title = state.course && state.step !== "import" ? `${STEPS.find((x) => x.key === state.step).label} · ${state.course} — Grading Console` : "Grading Console";
  }

  function renderSteps() {
    const ol = $("#steps");
    ol.replaceChildren(...STEPS.map((st) => {
      const avail = stepAvailable(st.key);
      const info = stepState(st.key);
      return h("li", { class: "step" + (avail ? "" : " is-locked"), "aria-current": state.step === st.key ? "step" : null },
        h("button", { type: "button", "aria-disabled": avail ? null : "true", onclick: () => go(st.key), title: avail ? null : "Import a file with at least one valid record first" },
          h("span", { class: "step__n" }, st.n),
          h("span", { class: "step__label" }, st.label),
          info ? h("span", { class: "step__state " + (info.cls || "") }, info.text) : null));
    }));
  }

  function stepState(key) {
    if (key === "import") {
      if (state.load) return { text: "Reading…" };
      if (state.fileError) return { text: "Needs a file", cls: "state-err" };
      if (hasData()) return { text: "✓ " + plural(state.parsed.counts.valid, "record"), cls: "state-ok" };
      return null;
    }
    if (!state.course || !hasData()) return null;
    const v = courseView(state.course);
    if (key === "review") return v.attention.length ? { text: `${v.attention.length} to check`, cls: "state-warn" } : { text: "✓ Clean", cls: "state-ok" };
    if (key === "grade") {
      if (!v.valid.length) return null;
      return C.validateBands(bandsFor(state.course)).length ? { text: "Ranges invalid", cls: "state-err" } : { text: "✓ Ranges valid", cls: "state-ok" };
    }
    if (key === "export") {
      const e = exportState(state.course);
      if (!e) return null;
      return e.stale ? { text: "Changed since export", cls: "state-warn" } : { text: "✓ Exported", cls: "state-ok" };
    }
  }

  function renderFileChip() {
    const chip = $("#fileChip");
    if (!state.file || !hasData()) { chip.hidden = true; return; }
    chip.hidden = false;
    chip.replaceChildren(
      h("span", { class: "fname", title: state.file.name }, state.file.name),
      h("span", { class: "muted" }, `· ${plural(state.parsed.courses.length, "course")}`),
      h("button", { class: "btn btn--quiet", type: "button", onclick: () => { state.step = "import"; render(); $("#fileInput") && $("#fileInput").click(); } }, "Replace"));
  }

  function renderRail() {
    const rail = $("#rail"), layout = $("#layout");
    const show = hasData() && state.step !== "import";
    rail.hidden = !show;
    layout.classList.toggle("has-rail", show);
    if (!show) return;
    const courses = state.parsed.courses;
    rail.replaceChildren(
      h("div", { class: "rail__head" },
        h("span", { class: "eyebrow" }, "Courses"),
        h("span", { class: "small muted" }, String(courses.length))),
      h("label", { class: "rail__select" },
        h("span", { class: "sr-only", hidden: true }, "Course"),
        h("select", { class: "select", "aria-label": "Course", onchange: (e) => selectCourse(e.target.value) },
          courses.map((c) => h("option", { value: c.name, selected: c.name === state.course }, `${c.name} (${c.valid})`)))),
      h("ul", { class: "rail__list" }, courses.map((c) => {
        const e = exportState(c.name);
        const bad = C.validateBands(bandsFor(c.name)).length > 0;
        return h("li", { class: "rail__item", "aria-current": c.name === state.course ? "true" : null },
          h("button", { type: "button", onclick: () => selectCourse(c.name), title: c.name },
            h("span", { class: "rail__name" }, c.name),
            h("span", { class: "rail__count num" }, c.valid.toLocaleString()),
            h("span", { class: "rail__meta" },
              c.valid === 0 ? h("span", {}, h("span", { class: "dot dot--err" }), " No gradeable records")
                : c.attention ? h("span", {}, h("span", { class: "dot dot--warn" }), ` ${c.attention} to check`)
                : h("span", {}, h("span", { class: "dot dot--ok" }), " Clean"),
              bad ? h("span", { class: "ico-err" }, "· ranges invalid") : null,
              e ? h("span", { class: e.stale ? "ico-warn" : "ico-ok" }, e.stale ? "· changed" : "· exported") : null)));
      })));
  }

  /* ------------------------------------------------------------------ *
   * STEP 01 — IMPORT
   * ------------------------------------------------------------------ */
  function viewImport() {
    const loaded = hasData() || (state.parsed && state.parsed.ok);
    return h("section", { "aria-labelledby": "t-import" },
      h("div", { class: "head" },
        h("div", {},
          h("div", { class: "eyebrow" }, "Step 01 · Import"),
          h("h1", { class: "display", id: "t-import" }, "Bring in the marks."),
          h("p", { class: "lede" }, "One Excel workbook, any number of courses. The file stays on this computer — nothing is uploaded."))),
      state.fileError ? importError() : null,
      loaded && !state.load ? importResult() : null,
      h("div", { class: "grid-2", style: loaded || state.fileError ? "margin-top:16px" : null },
        dropzone(loaded),
        formatSpec()));
  }

  function dropzone(loaded) {
    const input = h("input", {
      type: "file", id: "fileInput", accept: ".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel",
      onchange: (e) => { const f = e.target.files[0]; e.target.value = ""; handleFile(f); }, // reset so the same file can be re-imported
    });
    const zone = h("label", { class: "drop", for: "fileInput" },
      input,
      state.load
        ? h("div", { role: "status", "aria-live": "polite" },
            h("div", { class: "drop__title" }, state.load.phase === "reading" ? "Reading file…" : "Checking every row…"),
            h("div", { class: "small muted", style: "margin-top:4px" }, state.load.name),
            h("div", { class: "progress" }, h("span", { id: "loadBar", style: `width:${state.load.pct}%` })))
        : h("div", {},
            h("div", { class: "drop__icon", "aria-hidden": "true" }),
            h("div", { class: "drop__title", style: "margin-top:12px" }, loaded ? "Drop a corrected file to replace" : "Drop the marks file here"),
            h("div", { class: "muted", style: "margin-top:4px" }, "or ", h("span", { style: "color:var(--navy);text-decoration:underline" }, "choose a file"), " · .xlsx up to 25 MB"),
            loaded ? h("div", { class: "small muted", style: "margin-top:10px" }, "Grade ranges you’ve already set are kept for matching course names.") : null));
    ["dragenter", "dragover"].forEach((t) => zone.addEventListener(t, (e) => { e.preventDefault(); zone.classList.add("is-over"); }));
    ["dragleave", "drop"].forEach((t) => zone.addEventListener(t, () => zone.classList.remove("is-over")));
    zone.addEventListener("drop", (e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (e.dataTransfer.files.length > 1) toast("Only the first file was imported."); handleFile(f); });
    return h("div", {}, zone,
      !loaded ? h("div", { class: "actions", style: "margin-top:12px" },
        h("button", { class: "btn btn--quiet", type: "button", onclick: loadSample }, "Try with synthetic sample data"),
        h("button", { class: "btn btn--quiet", type: "button", onclick: downloadTemplate }, "Download blank template")) : null);
  }

  function formatSpec() {
    return h("div", { class: "panel" },
      h("div", { class: "panel__head" }, h("h2", { class: "section-title" }, "Expected format"), h("span", { class: "small muted" }, "First sheet with these headers")),
      h("table", { class: "spec" },
        h("thead", {}, h("tr", {}, h("th", {}, "Student’s BITS ID"), h("th", {}, "Course"), h("th", { class: "num" }, "Total Marks"))),
        h("tbody", {},
          [["2024A7PS0001P", "Course A", 82], ["2024A7PS0002P", "Course A", 71], ["2024A7PS0003P", "Course B", 64]]
            .map((r) => h("tr", {}, h("td", { class: "mono" }, r[0]), h("td", {}, r[1]), h("td", { class: "num" }, r[2]))))),
      h("ul", { class: "rules" },
        h("li", {}, "Total Marks are whole numbers from 0 to 100."),
        h("li", {}, "One row per student per course. A student may appear in several courses."),
        h("li", {}, "Leave out students who should receive NC — they did not appear for the examination."),
        h("li", {}, "Column order doesn’t matter; header capitalisation and apostrophe style are ignored.")));
  }

  function importError() {
    const e = state.fileError;
    return h("div", { class: "callout callout--err", role: "alert" },
      h("div", { class: "callout__title" }, e.title),
      h("p", {}, e.detail),
      e.found && e.found.length ? h("p", { class: "small", style: "margin-top:6px" }, "Headers found: ", e.found.map((f, i) => [i ? ", " : "", h("span", { class: "mono" }, f)])) : null,
      hasData() ? h("p", { class: "small muted", style: "margin-top:6px" }, `Your previous import (${state.file.name}) is still loaded.`) : null);
  }

  function importResult() {
    const p = state.parsed;
    const noCourse = p.records.filter((r) => !r.course);
    const courses = p.courses;
    if (!p.records.length) {
      return h("div", { class: "callout callout--warn", role: "status" },
        h("div", { class: "callout__title" }, "The headers are right, but there are no student rows"),
        h("p", {}, `Sheet “${state.file.sheet}” in ${state.file.name} has the three columns and nothing beneath them.`));
    }
    const notes = [];
    if (p.notes.textNumbers) notes.push(`${plural(p.notes.textNumbers, "mark")} stored as text in Excel ${p.notes.textNumbers === 1 ? "was" : "were"} read as ${p.notes.textNumbers === 1 ? "a number" : "numbers"}.`);
    if (p.notes.whitespaceFixed) notes.push(`${plural(p.notes.whitespaceFixed, "course name")} had stray spaces, which were removed so they group correctly.`);
    if (p.columns.extras.length) notes.push(`Ignored extra column${p.columns.extras.length > 1 ? "s" : ""}: ${p.columns.extras.join(", ")}.`);
    if (state.file.sheetCount > 1) notes.push(`Read sheet “${state.file.sheet}” (the workbook has ${state.file.sheetCount} sheets).`);
    if (p.counts.duplicates) notes.push(`${plural(p.counts.duplicates, "exact duplicate row")} will be counted once.`);

    const errs = p.counts.attention;
    return h("div", { class: "panel", role: "status" },
      h("div", { class: "panel__head" },
        h("div", {},
          h("div", { class: "eyebrow" }, errs ? "Imported — some records need attention" : "Marks imported"),
          h("div", { class: "result__line num" },
            plural(p.counts.valid, "student record"), h("span", { class: "sep" }, "·"),
            plural(courses.length, "course"), h("span", { class: "sep" }, "·"),
            h("span", { style: errs ? "color:var(--amber)" : "color:var(--sage)" }, errs ? `${errs.toLocaleString()} to check` : "0 errors"))),
        h("div", { class: "actions" }, h("button", { class: "btn btn--primary btn--lg", type: "button", onclick: () => go("review") }, "Review data →"))),
      h("p", { class: "small muted" }, `${state.file.name} · ${fmtBytes(state.file.size)} · imported ${timeOf(state.importedAt)}`),
      notes.length ? h("ul", { class: "rules" }, notes.map((n) => h("li", {}, n))) : null,
      h("hr", { class: "rule" }),
      h("div", { class: "table-wrap" },
        h("table", { class: "table" },
          h("thead", {}, h("tr", {}, h("th", {}, "Course"), h("th", { class: "num" }, "Valid"), h("th", { class: "num" }, "To check"), h("th", {}, ""))),
          h("tbody", {}, courses.map((c) => h("tr", {},
            h("td", {}, c.name),
            h("td", { class: "num" }, c.valid.toLocaleString()),
            h("td", { class: "num" }, c.attention ? h("span", { class: "tag tag--warn" }, c.attention) : h("span", { class: "muted" }, "—")),
            h("td", { class: "num" }, h("button", { class: "btn btn--quiet", type: "button", onclick: () => { state.course = c.name; go("review"); } }, "Open")))))),
        ),
      noCourse.length ? h("div", { class: "callout callout--warn", style: "margin-top:16px" },
        h("div", { class: "callout__title" }, `${plural(noCourse.length, "row")} ${noCourse.length === 1 ? "has" : "have"} no course and can’t be graded`),
        h("p", {}, "Add the course name in the file and import it again."),
        issueTable(noCourse)) : null);
  }

  /* ------------------------------------------------------------------ *
   * STEP 02 — REVIEW
   * ------------------------------------------------------------------ */
  function viewReview() {
    const v = courseView(state.course);
    const st = v.stats;
    const conflictIds = new Set(v.attention.filter((r) => r.issues.includes("duplicate-conflict")).map((r) => r.id.toUpperCase()));
    const missing = v.attention.filter((r) => r.issues.some((i) => i === "blank" || i === "missing-id")).length;
    const invalid = v.attention.filter((r) => r.issues.some((i) => i === "non-numeric" || i === "out-of-range" || i === "decimal")).length;
    const decimalsInFile = state.parsed.counts.decimals + (state.roundDecimals ? state.parsed.notes.rounded.length : 0);

    return h("section", { "aria-labelledby": "t-review" },
      h("div", { class: "head" },
        h("div", {},
          h("div", { class: "eyebrow" }, "Step 02 · Review"),
          h("h1", { class: "display", id: "t-review" }, state.course),
          h("p", { class: "lede" }, "Make sure the data is trustworthy before grading it.")),
        h("div", { class: "actions" },
          h("button", { class: "btn btn--primary btn--lg", type: "button", disabled: !v.valid.length, onclick: () => go("grade") }, "Continue to grading →"))),

      v.valid.length ? metrics(st) : h("div", { class: "callout callout--err" },
        h("div", { class: "callout__title" }, "No gradeable records in this course"),
        h("p", {}, "Every row for this course needs attention. Fix them in the file and import it again.")),

      h("div", { class: "quality", style: "margin-top:16px", "aria-label": "Data quality" },
        qualityCell(v.valid.length, "Valid records", false, true),
        qualityCell(v.attention.length, "Records requiring attention", v.attention.length > 0),
        qualityCell(conflictIds.size + v.duplicates.length, "Duplicate IDs", conflictIds.size > 0, false, v.duplicates.length && !conflictIds.size ? `${v.duplicates.length} exact repeat${v.duplicates.length > 1 ? "s" : ""}, counted once` : null),
        qualityCell(missing + invalid, "Missing or invalid marks", missing + invalid > 0)),

      v.attention.length ? attentionPanel(v, decimalsInFile) : h("div", { class: "callout callout--ok", style: "margin-top:16px" },
        h("div", { class: "callout__title" }, "Every record in this course is valid"),
        h("p", {}, v.duplicates.length ? `${plural(v.duplicates.length, "exact duplicate row")} ${v.duplicates.length === 1 ? "was" : "were"} ignored.` : "No missing values, invalid marks or duplicate IDs.")),

      state.roundDecimals && state.parsed.notes.rounded.some((r) => r.course === state.course) ? roundedPanel() : null,

      v.valid.length ? h("div", { class: "grid-2", style: "margin-top:16px" },
        h("div", { class: "panel" },
          h("div", { class: "panel__head" }, h("h2", { class: "section-title" }, "Distribution of marks"), h("span", { class: "small muted" }, "Bins of 10 marks")),
          decadeChart(v.marks, st),
          h("div", { class: "legend" }, h("span", {}, h("i"), "Mean"), h("span", {}, h("i", { style: "border-top-style:solid" }), "Median"))),
        h("div", { class: "panel" },
          h("div", { class: "panel__head" }, h("h2", { class: "section-title" }, "What the data says")),
          h("ul", { class: "insights" }, C.insights(v.marks, bandsFor(state.course)).map((t) => h("li", {}, t))))) : null,

      v.valid.length ? validTable(v) : null);
  }

  function metrics(st) {
    const m = (label, value, sub) => h("div", { class: "metric" },
      h("div", { class: "metric__label" }, label), h("div", { class: "metric__value num" }, value), sub ? h("div", { class: "metric__sub" }, sub) : null);
    return h("div", { class: "metrics", "aria-label": "Summary statistics" },
      m("Students", st.n.toLocaleString()),
      m("Mean", fmt(st.mean)),
      m("Median", fmt(st.median)),
      m("Highest", fmt(st.max)),
      m("Lowest", fmt(st.min)),
      m("Std. deviation", fmt(st.sd)));
  }

  function qualityCell(n, label, bad, good, sub) {
    return h("div", { class: bad ? "is-bad" : good ? "is-good" : "" },
      h("div", { class: "q__n num" }, n.toLocaleString()),
      h("div", { class: "q__l" }, label),
      sub ? h("div", { class: "q__l" }, sub) : null);
  }

  function attentionPanel(v, decimalsInFile) {
    const hasDecimals = v.attention.some((r) => r.issues.includes("decimal"));
    return h("div", { class: "panel", style: "margin-top:16px" },
      h("div", { class: "panel__head" },
        h("div", {},
          h("h2", { class: "section-title" }, "Records requiring attention"),
          h("p", { class: "small muted", style: "margin-top:2px" },
            `These ${plural(v.attention.length, "row")} ${v.attention.length === 1 ? "is" : "are"} excluded from grading and export. Fix them in the Excel file and drop it in again — your grade ranges are kept.`)),
        h("button", { class: "btn", type: "button", onclick: () => { state.step = "import"; render(); $("#fileInput").click(); } }, "Import corrected file")),
      hasDecimals && decimalsInFile ? h("div", { class: "callout callout--warn", style: "margin-bottom:12px" },
        h("div", { class: "callout__title" }, "Some marks aren’t whole numbers"),
        h("p", {}, "Marks must be whole numbers. A mark like 79.5 falls between two grades, so it can’t be graded as-is."),
        h("div", { class: "actions", style: "margin-top:10px" },
          h("button", { class: "btn", type: "button", onclick: () => { state.roundDecimals = true; reparse(); render(); toast(`${plural(state.parsed.notes.rounded.length, "mark")} rounded half-up.`, "Undo", () => { state.roundDecimals = false; reparse(); render(); }); } },
            `Round ${plural(decimalsInFile, "decimal mark")} half-up (79.5 → 80)`),
          h("span", { class: "small muted" }, "Applies to the whole file. Every change is listed so you can check it."))) : null,
      issueTable(v.attention));
  }

  function roundedPanel() {
    const list = state.parsed.notes.rounded.filter((r) => r.course === state.course);
    return h("details", { class: "panel disclose", style: "margin-top:16px" },
      h("summary", {}, `${plural(list.length, "mark")} rounded half-up in this course`),
      h("div", { class: "table-wrap", style: "margin-top:12px" }, h("table", { class: "table" },
        h("thead", {}, h("tr", {}, h("th", { class: "num" }, "Row"), h("th", {}, "BITS ID"), h("th", { class: "num" }, "In file"), h("th", { class: "num" }, "Used"))),
        h("tbody", {}, list.map((r) => h("tr", {}, h("td", { class: "num" }, r.row), h("td", { class: "mono" }, r.id), h("td", { class: "num" }, r.from), h("td", { class: "num" }, r.to)))))),
      h("div", { class: "actions", style: "margin-top:10px" },
        h("button", { class: "btn btn--quiet", type: "button", onclick: () => { state.roundDecimals = false; reparse(); render(); } }, "Undo rounding")));
  }

  function rawDisplay(v) {
    if (v === null || v === undefined || (typeof v === "string" && !v.trim())) return h("span", { class: "muted" }, "blank");
    return typeof v === "string" ? `“${v}”` : String(v);
  }

  function issueTable(rows) {
    return h("div", { class: "table-wrap table--scroll" },
      h("table", { class: "table" },
        h("thead", {}, h("tr", {}, h("th", { class: "num" }, "Row"), h("th", {}, "BITS ID"), h("th", { class: "num" }, "Total Marks"), h("th", {}, "Issue"))),
        h("tbody", {}, rows.map((r) => h("tr", {},
          h("td", { class: "num muted" }, r.row),
          h("td", { class: "mono" }, r.id || h("span", { class: "muted" }, "blank")),
          h("td", { class: "num" }, rawDisplay(r.raw)),
          h("td", { class: "wrap-ok" }, r.issues.map((i) => C.ISSUE_TEXT[i]).join("; ")))))));
  }

  function validTable(v) {
    const LIMIT = 300;
    const tbody = h("tbody");
    const fill = () => {
      const qq = state.reviewFilter.trim().toUpperCase();
      const list = qq ? v.valid.filter((r) => r.id.toUpperCase().includes(qq)) : v.valid;
      tbody.replaceChildren(...list.slice(0, LIMIT).map((r) => h("tr", {},
        h("td", { class: "num muted" }, r.row), h("td", { class: "mono" }, r.id), h("td", { class: "num" }, r.mark))));
      countEl.textContent = list.length > LIMIT ? `Showing ${LIMIT} of ${list.length.toLocaleString()} — search to narrow` : `${list.length.toLocaleString()} shown`;
    };
    const countEl = h("span", { class: "small muted" });
    const d = h("details", { class: "panel disclose", style: "margin-top:16px" },
      h("summary", {}, `Valid records (${v.valid.length.toLocaleString()})`),
      h("div", { style: "margin-top:12px;display:flex;gap:12px;align-items:center;flex-wrap:wrap" },
        h("input", { class: "input", type: "search", placeholder: "Search BITS ID", value: state.reviewFilter, style: "max-width:260px", "aria-label": "Search BITS ID",
          oninput: (e) => { state.reviewFilter = e.target.value; fill(); } }),
        countEl),
      h("div", { class: "table-wrap table--scroll", style: "margin-top:10px" },
        h("table", { class: "table" }, h("thead", {}, h("tr", {}, h("th", { class: "num" }, "Row"), h("th", {}, "BITS ID"), h("th", { class: "num" }, "Total Marks"))), tbody)));
    fill();
    return d;
  }

  /* ------------------------------------------------------------------ *
   * Charts (SVG)
   * ------------------------------------------------------------------ */
  function decadeChart(marks, st) {
    const W = 560, H = 250, P = { l: 34, r: 8, t: 34, b: 34 };
    const bins = C.histogram(marks);
    const maxN = Math.max(1, ...bins.map((b) => b.n));
    const ticks = niceTicks(maxN * 1.15); // headroom so value labels never touch the marker labels
    const top = ticks[ticks.length - 1];
    const x = (m) => P.l + (m / 101) * (W - P.l - P.r);
    const y = (n) => H - P.b - (n / top) * (H - P.t - P.b);
    const svg = s("svg", { class: "chart", viewBox: `0 0 ${W} ${H}`, role: "img", "aria-label": `Histogram of ${marks.length} marks in bins of ten. ` + bins.map((b) => `${b.lo}–${b.hi}: ${b.n}`).join(", ") });
    ticks.forEach((t) => {
      svg.appendChild(s("line", { class: "grid", x1: P.l, x2: W - P.r, y1: y(t), y2: y(t) }));
      svg.appendChild(s("text", { x: P.l - 8, y: y(t) + 4, "text-anchor": "end" }, t));
    });
    bins.forEach((b) => {
      const x0 = x(b.lo) + 2, x1 = x(b.hi + 1) - 2;
      const r = s("rect", { class: "bar", x: x0, width: Math.max(0, x1 - x0), y: y(b.n), height: H - P.b - y(b.n) },
        s("title", {}, `${b.lo}–${b.hi}: ${plural(b.n, "student")}`));
      svg.appendChild(r);
      if (b.n) svg.appendChild(s("text", { class: "val", x: (x0 + x1) / 2, y: y(b.n) - 5, "text-anchor": "middle" }, b.n));
      svg.appendChild(s("text", { x: (x0 + x1) / 2, y: H - P.b + 16, "text-anchor": "middle" }, `${b.lo}–${b.hi}`));
    });
    svg.appendChild(s("line", { class: "axis", x1: P.l, x2: W - P.r, y1: H - P.b, y2: H - P.b }));
    if (st.n) {
      const mx = x(st.mean + 0.5), dx = x(st.median + 0.5);
      svg.appendChild(s("line", { class: "marker", x1: mx, x2: mx, y1: 14, y2: H - P.b }));
      svg.appendChild(s("line", { class: "marker", x1: dx, x2: dx, y1: 14, y2: H - P.b, "stroke-dasharray": "none" }));
      const close = Math.abs(mx - dx) < 70;
      svg.appendChild(s("text", { class: "marker-label", x: mx + (close && mx < dx ? -4 : 4), y: 11, "text-anchor": close && mx < dx ? "end" : "start" }, `Mean ${fmt(st.mean)}`));
      svg.appendChild(s("text", { class: "marker-label", x: dx + (close && mx < dx ? 4 : -4), y: close ? 11 : 23, "text-anchor": close && mx < dx ? "start" : "end" }, `Median ${fmt(st.median)}`));
    }
    return svg;
  }

  function niceTicks(max) {
    const step = [1, 2, 5, 10, 20, 25, 50, 100, 200, 250, 500, 1000, 2000, 5000].find((s) => max / s <= 5) || Math.ceil(max / 5);
    const out = [];
    for (let t = 0; t <= max + step - 1 && out.length < 8; t += step) out.push(t);
    if (out[out.length - 1] < max) out.push(out[out.length - 1] + step);
    return out;
  }

  // Per-mark chart with grade bands shaded behind; the grading workspace's centrepiece.
  function bandChart(marks, bands, valid) {
    const W = 760, H = 210, P = { l: 30, r: 8, t: 26, b: 28 };
    const counts = Array(101).fill(0);
    marks.forEach((m) => counts[m]++);
    const maxN = Math.max(1, ...counts);
    const ticks = niceTicks(maxN);
    const top = ticks[ticks.length - 1];
    const x = (m) => P.l + (m / 101) * (W - P.l - P.r);
    const y = (n) => H - P.b - (n / top) * (H - P.t - P.b);
    const bw = (W - P.l - P.r) / 101;
    const svg = s("svg", { class: "chart", viewBox: `0 0 ${W} ${H}`, role: "img",
      "aria-label": "Marks by individual score with grade ranges shaded. " + (valid ? bands.map((b) => `${b.grade}: ${b.min}–${b.max}`).join(", ") : "Grade ranges are currently invalid.") });

    if (valid) {
      bands.forEach((b, i) => {
        svg.appendChild(s("rect", { class: "band-bg-" + (i % 2), x: x(b.min), y: P.t - 18, width: x(b.max + 1) - x(b.min), height: H - P.b - P.t + 18 }));
        const cx = (x(b.min) + x(b.max + 1)) / 2;
        if (x(b.max + 1) - x(b.min) >= 14) svg.appendChild(s("text", { class: "grade-label", x: cx, y: P.t - 6, "text-anchor": "middle" }, b.grade));
      });
    }
    ticks.forEach((t) => {
      svg.appendChild(s("line", { class: "grid", x1: P.l, x2: W - P.r, y1: y(t), y2: y(t), "stroke-opacity": t ? 0.7 : 0 }));
      svg.appendChild(s("text", { x: P.l - 6, y: y(t) + 4, "text-anchor": "end" }, t));
    });
    const bandIdx = (m) => bands.findIndex((b) => m >= b.min && m <= b.max);
    counts.forEach((n, m) => {
      if (!n) return;
      const bi = valid ? bandIdx(m) : 0;
      svg.appendChild(s("rect", { class: "band-bar-" + (bi % 2), x: x(m) + bw * 0.12, width: bw * 0.76, y: y(n), height: H - P.b - y(n) },
        s("title", {}, `Mark ${m}: ${plural(n, "student")}${valid ? " · " + bands[bi].grade : ""}`)));
    });
    if (valid) bands.slice(0, -1).forEach((b) => svg.appendChild(s("line", { class: "cut", x1: x(b.min), x2: x(b.min), y1: P.t - 18, y2: H - P.b })));
    svg.appendChild(s("line", { class: "axis", x1: P.l, x2: W - P.r, y1: H - P.b, y2: H - P.b }));
    for (let m = 0; m <= 100; m += 10) svg.appendChild(s("text", { x: x(m) + bw / 2, y: H - P.b + 16, "text-anchor": "middle" }, m));
    return svg;
  }

  /* ------------------------------------------------------------------ *
   * STEP 03 — GRADE
   * ------------------------------------------------------------------ */
  let lastBandsBeforeReset = null;

  function viewGrade() {
    const v = courseView(state.course);
    const bands = bandsFor(state.course);
    const rows = bands.map((b, i) => {
      const mk = (key) => h("input", {
        type: "number", inputmode: "numeric", min: 0, max: 100, step: 1, value: b[key], id: `band-${i}-${key}`,
        "aria-label": `${b.grade} ${key === "min" ? "lower" : "upper"} bound`,
        oninput: (e) => onBandInput(i, key, e.target.value),
        onblur: (e) => { if (e.target.value === "" && b[key] !== "") e.target.value = b[key]; },
      });
      return h("tr", { id: `band-row-${i}` },
        h("td", { class: "g" }, b.grade),
        h("td", {}, mk("min")),
        h("td", { class: "to", "aria-hidden": "true" }, "–"),
        h("td", {}, mk("max")),
        h("td", { class: "cnt num", id: `cnt-${i}` }),
        h("td", { class: "pct num", id: `pct-${i}` }),
        h("td", { class: "barcell" }, h("div", { class: "hbar", id: `bar-${i}` })));
    });

    return h("section", { "aria-labelledby": "t-grade" },
      h("div", { class: "head" },
        h("div", {},
          h("div", { class: "eyebrow" }, "Step 03 · Grade"),
          h("h1", { class: "display", id: "t-grade" }, state.course),
          h("p", { class: "lede" }, `Set the grade ranges. Every change updates the distribution for ${plural(v.valid.length, "student")} immediately.`)),
        h("div", { class: "actions" },
          h("button", { class: "btn btn--primary btn--lg", type: "button", id: "toExport", onclick: () => go("export") }, "Review & export →"))),

      h("div", { class: "panel" },
        h("div", { class: "panel__head" },
          h("h2", { class: "section-title" }, "Marks and grade ranges"),
          h("span", { class: "small muted" }, "Each bar is one mark; hover for counts")),
        h("div", { id: "bandChart" })),

      h("div", { class: "grid-2", style: "margin-top:16px" },
        h("div", { class: "panel" },
          h("div", { class: "panel__head" },
            h("h2", { class: "section-title" }, "Grade ranges"),
            h("label", { class: "switch", title: "When on, changing a bound moves the neighbouring grade’s bound so ranges stay contiguous." },
              h("input", { type: "checkbox", checked: state.linkBounds, onchange: (e) => { state.linkBounds = e.target.checked; } }),
              "Keep ranges contiguous")),
          h("div", { id: "bandStatus", "aria-live": "polite" }),
          h("table", { class: "bands", style: "margin-top:14px" },
            h("thead", {}, h("tr", {}, h("th", {}, "Grade"), h("th", {}, "From"), h("th", {}), h("th", {}, "To"),
              h("th", { class: "num", style: "text-align:right" }, "Students"), h("th", { style: "text-align:right" }, "%"), h("th", { class: "barcell" }, "Distribution"))),
            h("tbody", {}, rows)),
          h("div", { class: "actions", style: "margin-top:14px" },
            h("button", { class: "btn", type: "button", onclick: resetBands }, "Reset to default ranges"),
            copyFromControl())),
        h("div", { class: "stack" },
          h("div", { class: "panel" },
            h("div", { class: "panel__head" }, h("h2", { class: "section-title" }, "Close to a boundary")),
            h("div", { id: "nearBoundary" })),
          h("div", { class: "panel" },
            h("div", { class: "panel__head" }, h("h2", { class: "section-title" }, "Observations")),
            h("ul", { class: "insights", id: "gradeInsights" })))));
  }

  function copyFromControl() {
    const others = state.parsed.courses.filter((c) => c.name !== state.course && state.bandsByCourse.has(c.name)
      && JSON.stringify(state.bandsByCourse.get(c.name)) !== JSON.stringify(C.DEFAULT_BANDS)
      && !C.validateBands(state.bandsByCourse.get(c.name)).length);
    if (!others.length) return null;
    return h("select", { class: "select", style: "width:auto", "aria-label": "Copy ranges from another course",
      onchange: (e) => {
        if (!e.target.value) return;
        const prev = C.cloneBands(bandsFor(state.course));
        state.bandsByCourse.set(state.course, C.cloneBands(state.bandsByCourse.get(e.target.value)));
        render();
        toast(`Copied ranges from ${e.target.value}.`, "Undo", () => { state.bandsByCourse.set(state.course, prev); render(); });
      } },
      h("option", { value: "" }, "Copy ranges from…"),
      others.map((c) => h("option", { value: c.name }, c.name)));
  }

  function onBandInput(i, key, raw) {
    const bands = bandsFor(state.course);
    const val = raw === "" ? "" : Number(raw);
    bands[i][key] = val;
    if (state.linkBounds && Number.isInteger(val) && val >= 0 && val <= 100) {
      // Editing a lower bound moves the next grade's upper bound (and vice versa)
      // so no gap or overlap is created between neighbours.
      if (key === "min" && i < bands.length - 1 && val - 1 >= 0) {
        bands[i + 1].max = val - 1;
        const el = $(`#band-${i + 1}-max`); if (el) el.value = val - 1;
      }
      if (key === "max" && i > 0 && val + 1 <= 100) {
        bands[i - 1].min = val + 1;
        const el = $(`#band-${i - 1}-min`); if (el) el.value = val + 1;
      }
    }
    updateGrade();
    renderSteps();
    renderRail();
  }

  function resetBands() {
    const current = bandsFor(state.course);
    if (JSON.stringify(current) === JSON.stringify(C.DEFAULT_BANDS)) { toast("Ranges are already at the defaults."); return; }
    lastBandsBeforeReset = C.cloneBands(current);
    const course = state.course;
    state.bandsByCourse.set(course, C.cloneBands(C.DEFAULT_BANDS));
    render();
    toast("Grade ranges reset to defaults.", "Undo", () => {
      state.bandsByCourse.set(course, lastBandsBeforeReset);
      render();
    });
  }

  function updateGrade() {
    const v = courseView(state.course);
    const bands = bandsFor(state.course);
    const problems = C.validateBands(bands);
    const valid = problems.length === 0;
    const dist = C.gradeDistribution(v.marks, bands);

    // Status line
    const status = $("#bandStatus");
    if (status) {
      status.replaceChildren(valid
        ? h("div", { class: "status status--ok" }, h("span", { "aria-hidden": "true" }, "✓"),
            h("div", {}, h("strong", {}, "No gaps · No overlaps · 0–100 fully covered"),
              h("div", { class: "small", style: "color:var(--ink-2)" }, `All ${plural(v.valid.length, "student")} receive exactly one grade.`)))
        : h("div", { class: "status status--err", role: "alert" }, h("span", { "aria-hidden": "true" }, "✕"),
            h("div", {}, h("strong", {}, `${plural(problems.length, "problem")} with these ranges`),
              h("ul", {}, problems.map((p) => h("li", {}, p.message))),
              dist.unassigned ? h("div", { class: "small", style: "margin-top:6px;color:var(--ink)" }, `${plural(dist.unassigned, "student")} currently ${dist.unassigned === 1 ? "has" : "have"} no grade.`) : null)));
    }

    // Row states
    const badGrades = new Set(problems.flatMap((p) => p.grades));
    const badInputs = new Set();
    problems.forEach((p) => {
      const m = p.message.match(/^(\S+): (lower|upper) bound/);
      if (m) badInputs.add(`${m[1]}-${m[2] === "lower" ? "min" : "max"}`);
      if (p.type === "reversed") { badInputs.add(p.grades[0] + "-min"); badInputs.add(p.grades[0] + "-max"); }
      if (p.type === "overlap" && p.grades.length === 2) { badInputs.add(p.grades[0] + "-min"); badInputs.add(p.grades[1] + "-max"); }
      if (p.type === "order") p.grades.forEach((g) => { badInputs.add(g + "-min"); badInputs.add(g + "-max"); });
    });
    if (problems.some((p) => p.type === "gap" && /top of the scale/.test(p.message))) badInputs.add("A-max");
    if (problems.some((p) => p.type === "gap" && /bottom of the scale/.test(p.message))) badInputs.add("E-min");
    problems.filter((p) => p.type === "gap" && /between (\S+) and (\S+)\)/.test(p.message)).forEach((p) => {
      const [, lo, hi] = p.message.match(/between (\S+) and (\S+)\)/);
      badInputs.add(lo + "-max"); badInputs.add(hi + "-min");
    });

    const n = v.marks.length || 1;
    const maxCount = Math.max(1, ...dist.bands.map((b) => b.n));
    bands.forEach((b, i) => {
      const row = $(`#band-row-${i}`);
      if (!row) return;
      row.classList.toggle("is-bad", badGrades.has(b.grade));
      ["min", "max"].forEach((k) => {
        const el = $(`#band-${i}-${k}`);
        const bad = badInputs.has(`${b.grade}-${k}`);
        el.classList.toggle("is-invalid", bad);
        el.setAttribute("aria-invalid", bad ? "true" : "false");
      });
      const c = dist.bands[i].n;
      $(`#cnt-${i}`).textContent = c.toLocaleString();
      $(`#pct-${i}`).textContent = Math.round((c / n) * 100) + "%";
      const bar = $(`#bar-${i}`);
      bar.style.width = (c / maxCount) * 100 + "%";
      bar.classList.toggle("is-zero", c === 0);
      bar.title = `${b.grade}: ${plural(c, "student")}`;
    });

    const chart = $("#bandChart");
    if (chart) chart.replaceChildren(bandChart(v.marks, bands, valid));

    const nb = $("#nearBoundary");
    if (nb) {
      if (!valid) nb.replaceChildren(h("p", { class: "muted" }, "Shown once the ranges are valid."));
      else {
        const near = C.nearBoundary(v.valid, bands, 2);
        const one = near.filter((x) => x.gap === 1).length;
        nb.replaceChildren(
          near.length
            ? h("p", { class: "small muted", style: "margin-bottom:10px" },
                `${plural(near.length, "student")} ${near.length === 1 ? "is" : "are"} within 2 marks of the next grade up${one ? ` (${one} just 1 mark away)` : ""}.`)
            : h("p", { class: "muted" }, "No student is within 2 marks of the next grade up."),
          near.length ? h("div", { class: "table-wrap table--scroll", style: "max-height:240px" },
            h("table", { class: "table" },
              h("thead", {}, h("tr", {}, h("th", {}, "BITS ID"), h("th", { class: "num" }, "Mark"), h("th", {}, "Grade"), h("th", {}, "Next grade at"))),
              h("tbody", {}, near.map((x) => h("tr", {},
                h("td", { class: "mono" }, x.id), h("td", { class: "num" }, x.mark), h("td", {}, x.grade),
                h("td", {}, `${x.next} ≥ ${x.cutoff}`, h("span", { class: "muted" }, ` (+${x.gap})`))))))) : null);
      }
    }
    const ins = $("#gradeInsights");
    if (ins) ins.replaceChildren(...C.insights(v.marks, bands).map((t) => h("li", {}, t)));
    const btn = $("#toExport");
    if (btn) btn.textContent = valid ? "Review & export →" : "Review & export →";
  }

  /* ------------------------------------------------------------------ *
   * STEP 04 — EXPORT
   * ------------------------------------------------------------------ */
  function viewExport() {
    const v = courseView(state.course);
    const bands = bandsFor(state.course);
    const problems = C.validateBands(bands);
    const gaps = problems.filter((p) => p.type === "gap");
    const overlaps = problems.filter((p) => p.type === "overlap");
    const other = problems.filter((p) => p.type !== "gap" && p.type !== "overlap");
    const dist = C.gradeDistribution(v.marks, bands);
    const excluded = v.attention.length;
    const needsAck = excluded > 0;
    const acked = state.ack.has(state.course);
    const name = state.instructor.trim();
    const exp = exportState(state.course);

    const blockers = [];
    if (problems.length) blockers.push("fix the grade ranges");
    if (dist.unassigned) blockers.push("give every student a grade");
    if (!name) blockers.push("enter the instructor name");
    if (needsAck && !acked) blockers.push(`confirm the ${plural(excluded, "excluded record")}`);
    const ready = blockers.length === 0;

    const row = (ico, k, val, cls) => h("li", {},
      h("span", { class: "ico " + ({ ok: "ico-ok", warn: "ico-warn", err: "ico-err" }[cls] || "ico-info"), "aria-hidden": "true" }, ico),
      h("span", { class: "k" }, k), h("span", { class: "v num" }, val));

    const nameInput = h("input", { class: "input", id: "instructor", autocomplete: "name", placeholder: "e.g. Prof. A. Sharma", value: state.instructor,
      oninput: (e) => { state.instructor = e.target.value; storage.set("gc.instructor", e.target.value); refreshExportGate(); } });

    let csv = "";
    if (!problems.length && !dist.unassigned) {
      try { csv = C.buildCsv({ instructor: name || "(instructor)", course: state.course, records: v.valid, bands }); } catch (e) { csv = ""; }
    }
    const previewLines = csv.split("\r\n");
    const preview = previewLines.slice(0, 9).join("\n") + (previewLines.length > 10 ? `\n… ${v.valid.length - 5} more rows` : "");

    return h("section", { "aria-labelledby": "t-export" },
      h("div", { class: "head" },
        h("div", {},
          h("div", { class: "eyebrow" }, "Step 04 · Export"),
          h("h1", { class: "display", id: "t-export" }, state.course),
          h("p", { class: "lede" }, "A final check before the grades leave this console."))),

      exp ? h("div", { class: "callout " + (exp.stale ? "callout--warn" : "callout--ok"), style: "margin-bottom:16px", role: "status" },
        h("div", { class: "callout__title" }, exp.stale ? "Changed since your last export" : "Exported"),
        h("p", {}, exp.stale
          ? `You exported ${exp.filename} at ${timeOf(exp.at)}. The grade ranges or data have changed since then, so that file is out of date. Export again to replace it.`
          : `${plural(exp.count, "grade")} saved to ${exp.filename} at ${timeOf(exp.at)} — ${exp.elapsed}.`)) : null,

      h("div", { class: "grid-2" },
        h("div", { class: "panel" },
          h("div", { class: "panel__head" },
            h("span", { id: "readyLabel", class: "ready " + (ready ? "ready--ok" : "ready--no") }, ready ? "Ready to export" : "Not ready yet")),
          h("ul", { class: "checklist" },
            row("·", "Course", state.course),
            row("·", "Students to be graded", v.valid.length.toLocaleString()),
            row("·", "Grading bands", String(bands.length)),
            row(dist.unassigned ? "✕" : "✓", "Unassigned students", dist.unassigned.toLocaleString(), dist.unassigned ? "err" : "ok"),
            row(overlaps.length ? "✕" : "✓", "Overlapping ranges", overlaps.length ? String(overlaps.length) : "None", overlaps.length ? "err" : "ok"),
            row(gaps.length ? "✕" : "✓", "Gaps in 0–100", gaps.length ? String(gaps.length) : "None", gaps.length ? "err" : "ok"),
            other.length ? row("✕", "Other range problems", String(other.length), "err") : null,
            row(excluded ? "!" : "✓", "Records excluded (need attention)", excluded ? excluded.toLocaleString() : "None", excluded ? "warn" : "ok"),
            v.duplicates.length ? row("·", "Exact duplicate rows ignored", v.duplicates.length.toLocaleString()) : null),

          problems.length ? h("div", { class: "callout callout--err", style: "margin-top:14px" },
            h("div", { class: "callout__title" }, "The grade ranges need fixing"),
            h("ul", {}, problems.map((p) => h("li", {}, p.message))),
            h("div", { class: "actions", style: "margin-top:8px" }, h("button", { class: "btn", type: "button", onclick: () => go("grade") }, "Back to grade ranges"))) : null,

          h("div", { class: "field", style: "margin-top:18px" },
            h("label", { for: "instructor" }, "Instructor name ", h("span", { class: "muted" }, "(written into the file)")),
            nameInput),

          needsAck ? h("label", { class: "check", style: "margin-top:14px" },
            h("input", { type: "checkbox", id: "ackBox", checked: acked, onchange: (e) => { e.target.checked ? state.ack.add(state.course) : state.ack.delete(state.course); refreshExportGate(); } }),
            h("span", {}, `I understand ${plural(excluded, "record")} requiring attention will `, h("strong", {}, "not"), " appear in this export. ",
              h("button", { class: "btn btn--quiet", type: "button", style: "height:auto;padding:0", onclick: (e) => { e.preventDefault(); go("review"); } }, "See which"))) : null,

          h("div", { class: "actions", style: "margin-top:18px" },
            h("button", { class: "btn btn--primary btn--lg", type: "button", id: "exportBtn", disabled: !ready, onclick: doExport }, exp && !exp.stale ? "Export again" : "Export grades (.csv)"),
            h("span", { class: "small muted", id: "exportHint" }, ready ? C.exportFilename(state.course) : "To export: " + blockers.join(", ") + "."))),

        h("div", { class: "stack" },
          h("div", { class: "panel" },
            h("div", { class: "panel__head" }, h("h2", { class: "section-title" }, "Grade distribution")),
            problems.length ? h("p", { class: "muted" }, "Shown once the ranges are valid.") :
              h("table", { class: "table" },
                h("thead", {}, h("tr", {}, h("th", {}, "Grade"), h("th", {}, "Range"), h("th", { class: "num" }, "Students"), h("th", { class: "num" }, "%"))),
                h("tbody", {}, dist.bands.map((b) => h("tr", {},
                  h("td", { style: "font-weight:600" }, b.grade), h("td", { class: "num", style: "text-align:left" }, `${b.min}–${b.max}`),
                  h("td", { class: "num" }, b.n.toLocaleString()), h("td", { class: "num muted" }, Math.round((b.n / (v.marks.length || 1)) * 100) + "%")))))),
          csv ? h("div", { class: "panel" },
            h("div", { class: "panel__head" }, h("h2", { class: "section-title" }, "File preview"), h("span", { class: "small muted" }, "CSV · UTF-8")),
            h("pre", { class: "preview", id: "csvPreview" }, preview)) : null)));
  }

  // Updates the export gate in place so typing the name doesn't rebuild the form.
  function refreshExportGate() {
    const v = courseView(state.course);
    const bands = bandsFor(state.course);
    const problems = C.validateBands(bands);
    const dist = C.gradeDistribution(v.marks, bands);
    const blockers = [];
    if (problems.length) blockers.push("fix the grade ranges");
    if (dist.unassigned) blockers.push("give every student a grade");
    if (!state.instructor.trim()) blockers.push("enter the instructor name");
    if (v.attention.length && !state.ack.has(state.course)) blockers.push(`confirm the ${plural(v.attention.length, "excluded record")}`);
    const ready = !blockers.length;
    $("#exportBtn").disabled = !ready;
    $("#exportHint").textContent = ready ? C.exportFilename(state.course) : "To export: " + blockers.join(", ") + ".";
    const lab = $("#readyLabel");
    lab.textContent = ready ? "Ready to export" : "Not ready yet";
    lab.className = "ready " + (ready ? "ready--ok" : "ready--no");
    const pre = $("#csvPreview");
    if (pre) pre.textContent = pre.textContent.replace(/^Instructor,.*$/m, C.csvCell("Instructor") + "," + C.csvCell(state.instructor.trim() || "(instructor)"));
  }

  function doExport() {
    const v = courseView(state.course);
    const bands = bandsFor(state.course);
    let csv;
    try {
      csv = C.buildCsv({ instructor: state.instructor.trim(), course: state.course, records: v.valid, bands });
    } catch (err) {
      toast("Export stopped: " + err.message);
      return;
    }
    const filename = C.exportFilename(state.course);
    // BOM so Excel opens UTF-8 names (e.g. “Student’s”) correctly.
    const url = URL.createObjectURL(new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" }));
    const a = h("a", { href: url, download: filename, style: "display:none" });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);

    const now = new Date();
    const secs = Math.max(0, Math.round((now - state.importedAt) / 1000));
    const elapsed = secs < 60 ? `${secs} s after import` : `${Math.floor(secs / 60)} min ${secs % 60} s after import`;
    state.exports.set(state.course, { at: now, signature: signature(state.course), count: v.valid.length, filename, elapsed });
    render();
    const next = state.parsed.courses.find((c) => c.valid > 0 && !state.exports.has(c.name));
    if (next) toast(`Exported ${state.course}.`, `Next: ${next.name}`, () => { state.course = next.name; state.step = "review"; render(); });
    else toast(`Exported ${plural(v.valid.length, "grade")}.`);
  }

  /* ------------------------------------------------------------------ *
   * Toast
   * ------------------------------------------------------------------ */
  let toastTimer = null;
  function toast(msg, actionLabel, action) {
    const t = $("#toast");
    clearTimeout(toastTimer);
    t.replaceChildren(h("span", {}, msg), actionLabel ? h("button", { type: "button", onclick: () => { t.hidden = true; action(); } }, actionLabel) : null);
    t.hidden = false;
    toastTimer = setTimeout(() => { t.hidden = true; }, actionLabel ? 9000 : 4000);
  }

  /* ------------------------------------------------------------------ *
   * Boot
   * ------------------------------------------------------------------ */
  function boot() {
    // A file dropped anywhere outside the drop zone must not navigate away and lose work.
    window.addEventListener("dragover", (e) => e.preventDefault());
    window.addEventListener("drop", (e) => {
      e.preventDefault();
      if (e.dataTransfer && e.dataTransfer.files.length && !e.target.closest(".drop")) handleFile(e.dataTransfer.files[0]);
    });
    window.addEventListener("beforeunload", (e) => {
      if (hasData() && state.parsed.courses.some((c) => c.valid > 0 && !(exportState(c.name) && !exportState(c.name).stale))
          && [...state.bandsByCourse.values()].some((b) => JSON.stringify(b) !== JSON.stringify(C.DEFAULT_BANDS))) {
        e.preventDefault();
        e.returnValue = "";
      }
    });
    if (typeof XLSX === "undefined") {
      state.fileError = { title: "The spreadsheet reader didn’t load", detail: "Reload the page. If this keeps happening, the file vendor/xlsx.full.min.js is missing from the deployment." };
    }
    render();
  }

  // Expose a tiny surface for automated tests.
  window.__gc = { state, handleFile, render };
  boot();
})();
