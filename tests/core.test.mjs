// Unit tests for src/core.js. Run: node --test tests/
import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import * as XLSX from "xlsx";
import * as fs from "node:fs";

const require = createRequire(import.meta.url);
const C = require("../src/core.js");
const FIX = (f) => new URL("../fixtures/" + f, import.meta.url).pathname;
const H = ["Student’s BITS ID", "Course", "Total Marks"];
const sheet = (f) => {
  const wb = XLSX.read(fs.readFileSync(FIX(f)));
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null, blankrows: false });
};
const bands = () => C.cloneBands(C.DEFAULT_BANDS);
const find = (res, id, course) => res.records.filter((r) => r.id === id && (!course || r.course === course));

test("header detection tolerates apostrophe style, case and spacing", () => {
  for (const h of [H, ["Student's BITS ID", " course ", "TOTAL MARKS"], ["BITS ID", "Course", "Total"]]) {
    const r = C.parseSheet([h, ["2024A", "X", 50]]);
    assert.equal(r.ok, true, JSON.stringify(h));
    assert.equal(r.records[0].id, "2024A");
  }
});

test("column order does not matter", () => {
  const r = C.parseSheet([["Total Marks", "Course", "Student’s BITS ID"], [88, "X", "2024A"]]);
  assert.equal(r.records[0].mark, 88);
  assert.equal(r.records[0].id, "2024A");
});

test("wrong / missing columns are a hard, explained error", () => {
  const r = C.parseSheet(sheet("wrong-columns.xlsx"));
  assert.equal(r.ok, false);
  assert.equal(r.error.code, "columns");
  assert.deepEqual(r.error.missing, ["Student’s BITS ID", "Course", "Total Marks"]);
  assert.deepEqual(r.error.found, ["ID No", "Subject", "Marks"]);
  const m = C.parseSheet(sheet("missing-marks-column.xlsx"));
  assert.deepEqual(m.error.missing, ["Total Marks"]);
});

test("empty workbook and header-only workbook", () => {
  assert.equal(C.parseSheet(sheet("empty.xlsx")).error.code, "empty");
  const h = C.parseSheet(sheet("header-only.xlsx"));
  assert.equal(h.ok, true);
  assert.equal(h.records.length, 0);
  assert.equal(h.courses.length, 0);
});

test("happy path: 248 students, 3 courses, no issues", () => {
  const r = C.parseSheet(sheet("sample-marks.xlsx"));
  assert.equal(r.counts.rows, 248);
  assert.equal(r.counts.valid, 248);
  assert.equal(r.counts.attention, 0);
  assert.deepEqual(r.courses.map((c) => [c.name, c.valid]), [
    ["Data Structures & Algorithms", 120], ["Discrete Mathematics", 50], ["Operating Systems", 78],
  ]);
});

test("every data-quality problem is flagged, never graded silently", () => {
  const r = C.parseSheet(sheet("edge-cases.xlsx"));
  const issue = (id) => find(r, id).flatMap((x) => x.issues);
  assert.deepEqual(issue("2024A7PS0004P"), ["out-of-range"]);   // -5
  assert.deepEqual(issue("2024A7PS0005P"), ["out-of-range"]);   // 104
  assert.deepEqual(issue("2024A7PS0006P"), ["decimal"]);        // 79.5
  assert.deepEqual(issue("2024A7PS0007P"), ["blank"]);
  assert.deepEqual(issue("2024A7PS0008P"), ["non-numeric"]);    // "AB"
  assert.deepEqual(issue("2024A7PS0011P"), ["missing-course"]);
  assert.equal(find(r, "2024A7PS0009P")[0].mark, 68);            // "68" as text is accepted
  assert.equal(find(r, "2024A7PS0009P")[0].status, "valid");
  // Conflicting duplicate → both held back
  assert.deepEqual(find(r, "2024A7PS0002P").map((x) => x.status), ["attention", "attention"]);
  // Exact duplicate → second ignored, first graded
  assert.deepEqual(find(r, "2024A7PS0003P").map((x) => x.status), ["valid", "duplicate"]);
  // Missing ID
  assert.ok(r.records.some((x) => x.id === "" && x.issues.includes("missing-id")));
  // "Course A " merged into "Course A"
  assert.deepEqual(r.courses.map((c) => c.name), ["Course A", "Course B"]);
  assert.equal(r.notes.whitespaceFixed, 1);
  // Valid Course A: 0001(82) 0003(55) 0009(68) 0010(72) 0012(90)
  const valid = r.records.filter((x) => x.course === "Course A" && x.status === "valid").map((x) => x.mark);
  assert.deepEqual(valid, [82, 55, 68, 72, 90]);
});

test("decimal marks are rounded half-up only when the instructor opts in", () => {
  const r = C.parseSheet(sheet("edge-cases.xlsx"), { roundDecimals: true });
  const rec = find(r, "2024A7PS0006P")[0];
  assert.equal(rec.mark, 80);
  assert.equal(rec.status, "valid");
  assert.deepEqual(r.notes.rounded.map((x) => [x.from, x.to]), [[79.5, 80]]);
  assert.equal(C.parseSheet([H, ["a", "X", 80.2]], { roundDecimals: true }).records[0].mark, 80);
  assert.equal(C.parseSheet([H, ["a", "X", 80.49]], { roundDecimals: true }).records[0].mark, 80);
});

test("same BITS ID in two different courses is not a duplicate", () => {
  const r = C.parseSheet([H, ["a", "X", 50], ["a", "Y", 60]]);
  assert.equal(r.counts.valid, 2);
});

test("stats: numeric (not string) arithmetic, empty and single-value sets", () => {
  const r = C.parseSheet([H, ["a", "X", "70"], ["b", "X", "72"], ["c", "X", 50]]);
  const s = C.stats(r.records.map((x) => x.mark));
  assert.equal(s.mean, 64);
  assert.equal(s.median, 70);
  assert.equal(s.min, 50);
  assert.equal(s.max, 72);
  assert.deepEqual(C.stats([]), { n: 0, min: null, max: null, mean: null, median: null, sd: null });
  assert.equal(C.stats([75, 75]).sd, 0);
  assert.equal(C.stats([1, 2, 3, 4]).median, 2.5);
});

test("histogram bins are 0–9 … 90–100 and include both ends", () => {
  const h = C.histogram([0, 9, 10, 99, 100]);
  assert.equal(h.length, 10);
  assert.deepEqual([h[0].lo, h[0].hi, h[9].lo, h[9].hi], [0, 9, 90, 100]);
  assert.deepEqual(h.map((b) => b.n), [2, 1, 0, 0, 0, 0, 0, 0, 0, 2]);
});

test("default bands are valid and cover every mark exactly once", () => {
  assert.deepEqual(C.validateBands(bands()), []);
  for (let m = 0; m <= 100; m++) assert.ok(C.gradeFor(m, bands()), `mark ${m}`);
  assert.equal(C.gradeFor(80, bands()), "A");
  assert.equal(C.gradeFor(79, bands()), "A-");
  assert.equal(C.gradeFor(0, bands()), "E");
});

test("band validation explains gaps, overlaps, reversals, range and order", () => {
  const set = (g, k, v) => { const b = bands(); b.find((x) => x.grade === g)[k] = v; return b; };
  const msgs = (b) => C.validateBands(b).map((p) => p.type + ": " + p.message);

  assert.deepEqual(msgs(set("A", "max", 90)), ["gap: No grade covers marks 91–100 (top of the scale)."]);
  assert.deepEqual(msgs(set("E", "min", 5)), ["gap: No grade covers marks 0–4 (bottom of the scale)."]);
  assert.deepEqual(msgs(set("A-", "max", 75)), ["gap: No grade covers marks 76–79 (between A- and A)."]);
  assert.deepEqual(msgs(set("A-", "max", 82)), ["overlap: A and A- overlap on marks 80–82."]);
  assert.deepEqual(msgs(set("A", "max", 105)), ["range: A: upper bound 105 is outside 0–100."]);
  assert.deepEqual(msgs(set("B", "min", 72.5)), ["invalid: B: lower bound 72.5 must be a whole number."]);
  assert.deepEqual(msgs(set("B", "min", "")), ["invalid: B: lower bound is empty."]);
  const rev = msgs(set("B", "min", 69).map((x) => (x.grade === "B" ? { ...x, min: 69, max: 60 } : x)));
  assert.ok(rev[0].startsWith("reversed: B: lower bound 69 is above upper bound 60."), rev.join("\n"));
});

test("a single-mark band (min = max) is legitimate", () => {
  const b = bands();
  b[0] = { grade: "A", min: 100, max: 100 };
  b[1] = { grade: "A-", min: 70, max: 99 };
  assert.deepEqual(C.validateBands(b), []);
  assert.equal(C.gradeFor(100, b), "A");
  assert.equal(C.gradeFor(99, b), "A-");
});

test("inverted scheme (E on top) is rejected even with no gaps or overlaps", () => {
  const b = bands().reverse().map((x, i) => ({ grade: C.GRADES[i], min: x.min, max: x.max }));
  const p = C.validateBands(b);
  assert.ok(p.length && p.every((x) => x.type === "order"), JSON.stringify(p));
});

test("CSV export: escaping, formula neutralising, and original layout", () => {
  const csv = C.buildCsv({
    instructor: 'Rao, K. "Senior"',
    course: "Course, With \"Comma\"",
    records: [{ id: '=HYPERLINK("x")', mark: 77 }, { id: "2024A", mark: 19 }],
    bands: bands(),
  });
  assert.equal(csv,
    'Instructor,"Rao, K. ""Senior"""\r\n' +
    'Course,"Course, With ""Comma"""\r\n' +
    "\r\n" +
    "BITS ID,Total Marks,Grade\r\n" +
    '"\'=HYPERLINK(""x"")",77,A-\r\n' +
    "2024A,19,E\r\n");
});

test("CSV export refuses to run with invalid bands or an ungradeable student", () => {
  const b = bands(); b[0].max = 90;
  assert.throws(() => C.buildCsv({ instructor: "x", course: "y", records: [{ id: "a", mark: 95 }], bands: b }), /invalid/);
  assert.throws(() => C.buildCsv({ instructor: "x", course: "y", records: [{ id: "a", mark: 79.5 }], bands: bands() }), /no grade/);
});

test("export filename is safe and dated", () => {
  assert.equal(C.exportFilename("Data Structures & Algorithms", new Date(2026, 8, 30)), "grades_data-structures-algorithms_2026-09-30.csv");
  assert.equal(C.exportFilename("///", new Date(2026, 0, 2)), "grades_course_2026-01-02.csv");
});

test("near-boundary students are found", () => {
  const recs = [{ id: "a", mark: 79 }, { id: "b", mark: 78 }, { id: "c", mark: 80 }, { id: "d", mark: 69 }];
  assert.deepEqual(C.nearBoundary(recs, bands(), 1).map((x) => x.id), ["a", "d"]);
  assert.deepEqual(C.nearBoundary(recs, bands(), 2).map((x) => x.id), ["a", "d", "b"]);
});

test("insights are factual", () => {
  const marks = [62, 65, 68, 70, 71, 75, 78, 40, 90, 66];
  const ins = C.insights(marks, bands());
  assert.ok(ins[0].startsWith("80% of students (8 of 10) scored between 60 and 79."), ins[0]);
  assert.deepEqual(C.insights([75, 75, 75, 75, 75], bands()).slice(0, 1), ["Every student scored 75."]);
});

test("large dataset: 20,000 rows parse quickly", () => {
  const rows = sheet("large-20k.xlsx");
  const t = performance.now();
  const r = C.parseSheet(rows);
  const ms = performance.now() - t;
  assert.equal(r.counts.valid, 20000);
  assert.equal(r.courses.length, 40);
  assert.ok(ms < 1000, `parse took ${ms}ms`);
});
