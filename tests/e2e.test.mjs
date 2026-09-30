// End-to-end tests for the redesigned console, in headless Chromium.
// Run: node --test tests/e2e.test.mjs   (SCREENSHOTS=dir to save screenshots)
import test, { before, after } from "node:test";
import assert from "node:assert/strict";
import { chromium } from "playwright";
import { createServer } from "node:http";
import { readFileSync, existsSync, statSync, mkdirSync } from "node:fs";
import path from "node:path";

const ROOT = path.resolve(new URL("..", import.meta.url).pathname);
const FIX = (f) => path.join(ROOT, "fixtures", f);
const SHOTS = process.env.SCREENSHOTS;
if (SHOTS) mkdirSync(SHOTS, { recursive: true });
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".xlsx": "application/octet-stream" };

let server, browser, base;
before(async () => {
  server = createServer((req, res) => {
    const p = path.join(ROOT, decodeURIComponent(new URL(req.url, "http://x").pathname));
    const file = p.endsWith("/") ? p + "index.html" : p;
    if (!file.startsWith(ROOT) || !existsSync(file) || statSync(file).isDirectory()) { res.writeHead(404); return res.end(); }
    res.writeHead(200, { "content-type": TYPES[path.extname(file)] || "application/octet-stream" });
    res.end(readFileSync(file));
  });
  await new Promise((r) => server.listen(0, r));
  // BASE_URL=https://… runs the same suite against a deployment.
  base = process.env.BASE_URL || `http://127.0.0.1:${server.address().port}/`;
  browser = await chromium.launch();
});
after(async () => { await browser?.close(); server?.close(); });

async function open(viewport = { width: 1360, height: 900 }) {
  // ignoreHTTPSErrors only for BASE_URL runs: sandboxed CI egress may sit behind a TLS-inspecting proxy.
  const ctx = await browser.newContext({ viewport, acceptDownloads: true, ignoreHTTPSErrors: !!process.env.BASE_URL });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => m.type() === "error" && !/fonts\.(googleapis|gstatic)/.test(m.text()) && errors.push(m.text()));
  if (!process.env.BASE_URL) await page.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.fulfill({ body: "", contentType: "text/css" }));
  await page.goto(base);
  return { page, errors, close: () => ctx.close() };
}
const upload = async (page, f) => {
  await page.setInputFiles("#fileInput", FIX(f));
  await page.waitForFunction(() => !window.__gc.state.load, null, { timeout: 30000 });
};
const shot = async (page, name) => SHOTS && page.screenshot({ path: path.join(SHOTS, name + ".png"), fullPage: true });
const stepBtn = (page, label) => page.locator(".step button", { hasText: label });

test("happy path: import → review → grade → export", async () => {
  const { page, errors, close } = await open();
  await shot(page, "01-import-empty");
  await upload(page, "sample-marks.xlsx");
  const line = await page.textContent(".result__line");
  assert.match(line, /248 student records.*3 courses.*0 errors/);
  await shot(page, "02-import-done");

  await page.click("text=Review data →");
  assert.equal(await page.textContent("h1"), "Data Structures & Algorithms");
  const metric = async (label) => page.locator(".metric", { hasText: label }).locator(".metric__value").textContent();
  assert.equal(await metric("Students"), "120");
  const st = await page.evaluate(() => {
    const m = window.__gc.state.parsed.records.filter((r) => r.course === "Data Structures & Algorithms").map((r) => r.mark).sort((a, b) => a - b);
    return { min: m[0], max: m[m.length - 1] };
  });
  assert.equal(await metric("Lowest"), String(st.min));
  assert.equal(await metric("Highest"), String(st.max));
  await shot(page, "03-review");

  await page.click("text=Continue to grading →");
  assert.match(await page.textContent("#bandStatus"), /No gaps · No overlaps · 0–100 fully covered/);
  const total = await page.$$eval("[id^=cnt-]", (els) => els.reduce((a, e) => a + Number(e.textContent.replace(/,/g, "")), 0));
  assert.equal(total, 120);
  await shot(page, "04-grade");

  // Linked editing: raising A's lower bound to 85 moves A-'s upper bound to 84.
  await page.fill("#band-0-min", "85");
  assert.equal(await page.inputValue("#band-1-max"), "84");
  assert.match(await page.textContent("#bandStatus"), /No gaps/);

  await page.click("text=Review & export →");
  assert.equal(await page.isDisabled("#exportBtn"), true, "needs instructor name");
  assert.match(await page.textContent("#exportHint"), /enter the instructor name/);
  await page.fill("#instructor", "Prof. A. Sharma");
  assert.equal(await page.isDisabled("#exportBtn"), false);
  await shot(page, "05-export-ready");

  const [dl] = await Promise.all([page.waitForEvent("download"), page.click("#exportBtn")]);
  assert.match(dl.suggestedFilename(), /^grades_data-structures-algorithms_\d{4}-\d{2}-\d{2}\.csv$/);
  const csv = readFileSync(await dl.path(), "utf8").replace(/^﻿/, "");
  const lines = csv.trim().split("\r\n");
  assert.equal(lines[0], "Instructor,Prof. A. Sharma");
  assert.equal(lines[1], "Course,Data Structures & Algorithms");
  assert.equal(lines[3], "BITS ID,Total Marks,Grade");
  assert.equal(lines.length - 4, 120, "every student exported");
  assert.ok(lines.slice(4).every((l) => /^2024A7PS\d{4}P,\d{1,3},(A|A-|B|B-|C|C-|D|E)$/.test(l)), "no undefined IDs");
  // Cutoff change is honoured: 84 is A-, 85 is A.
  for (const l of lines.slice(4)) {
    const [, m, g] = l.split(",");
    if (+m >= 85) assert.equal(g, "A");
    if (+m === 84 || (+m >= 70 && +m < 85)) assert.equal(g, "A-");
  }
  assert.match(await page.textContent(".callout--ok"), /120 grades saved/);
  await shot(page, "06-exported");

  // Changing ranges after export is flagged.
  await stepBtn(page, "Grade").click();
  await page.fill("#band-2-min", "58");
  await stepBtn(page, "Export").click();
  assert.match(await page.textContent(".callout--warn"), /Changed since your last export/);
  assert.deepEqual(errors, []);
  await close();
});

test("gaps, overlaps, out-of-range and reversed bounds block export with precise messages", async () => {
  const { page, errors, close } = await open();
  await upload(page, "sample-marks.xlsx");
  await page.click("text=Review data →");
  await page.click("text=Continue to grading →");
  await page.uncheck("text=Keep ranges contiguous");

  await page.fill("#band-0-max", "90");
  assert.match(await page.textContent("#bandStatus"), /No grade covers marks 91–100 \(top of the scale\)/);
  assert.equal(await page.getAttribute("#band-0-max", "aria-invalid"), "true");
  await page.fill("#band-0-max", "100");

  await page.fill("#band-1-max", "83");
  assert.match(await page.textContent("#bandStatus"), /A and A- overlap on marks 80–83/);
  await shot(page, "07-grade-overlap");
  await stepBtn(page, "Export").click();
  await page.fill("#instructor", "X");
  assert.equal(await page.isDisabled("#exportBtn"), true);
  assert.match(await page.textContent(".checklist"), /Overlapping ranges\s*1/);
  await shot(page, "08-export-blocked");
  await stepBtn(page, "Grade").click();
  await page.fill("#band-1-max", "79");

  await page.fill("#band-7-min", "-3");
  assert.match(await page.textContent("#bandStatus"), /E: lower bound -3 is outside 0–100/);
  await page.fill("#band-7-min", "0");

  await page.fill("#band-2-min", "69");
  await page.fill("#band-2-max", "60");
  assert.match(await page.textContent("#bandStatus"), /B: lower bound 69 is above upper bound 60/);
  await page.click("text=Reset to default ranges");
  assert.match(await page.textContent("#bandStatus"), /No gaps/);
  // Undo restores the broken state (reset is reversible, no double confirm).
  await page.click("#toast button");
  assert.match(await page.textContent("#bandStatus"), /above upper bound/);
  assert.deepEqual(errors, []);
  await close();
});

test("edge-case file: every bad record surfaced; nothing silently dropped", async () => {
  const { page, errors, close } = await open();
  await upload(page, "edge-cases.xlsx");
  assert.match(await page.textContent(".result__line"), /6 student records.*2 courses.*9 to check/);
  assert.match(await page.textContent("main"), /1 row has no course and can’t be graded/);
  assert.match(await page.textContent("main"), /course name had stray spaces/);
  await shot(page, "09-import-edge");
  await page.click("text=Review data →");
  const table = await page.textContent(".panel:has(h2:text('Records requiring attention'))");
  for (const t of ["outside 0–100", "not a whole number", "Total Marks is blank", "not a number", "different mark", "BITS ID is blank"]) assert.match(table, new RegExp(t));
  await shot(page, "10-review-edge");

  // Round decimals: 79.5 → 80, listed, undoable.
  await page.click("text=/Round 1 decimal mark half-up/");
  assert.match(await page.textContent("main"), /1 mark rounded half-up in this course/);

  await page.click("text=Continue to grading →");
  await stepBtn(page, "Export").click();
  await page.fill("#instructor", "Dr. Test");
  assert.equal(await page.isDisabled("#exportBtn"), true, "needs acknowledgement of excluded records");
  await page.check("#ackBox");
  const [dl] = await Promise.all([page.waitForEvent("download"), page.click("#exportBtn")]);
  const rows = readFileSync(await dl.path(), "utf8").trim().split("\r\n").slice(4);
  assert.deepEqual(rows, [
    "2024A7PS0001P,82,A", "2024A7PS0003P,55,B-", "2024A7PS0006P,80,A",
    "2024A7PS0009P,68,B", "2024A7PS0010P,72,A-", "2024A7PS0012P,90,A",
  ]);
  assert.deepEqual(errors, []);
  await close();
});

test("bad files give a clear message and keep the previous import", async () => {
  const { page, errors, close } = await open();
  for (const [f, re] of [
    ["empty.xlsx", /The workbook is empty/],
    ["wrong-columns.xlsx", /required columns weren’t found[\s\S]*Missing required columns: Student’s BITS ID, Course, Total Marks[\s\S]*Headers found: ID No, Subject, Marks/],
    ["missing-marks-column.xlsx", /Missing required column: Total Marks/],
  ]) {
    await upload(page, f);
    assert.match(await page.textContent("[role=alert]"), re, f);
  }
  await upload(page, "header-only.xlsx");
  assert.match(await page.textContent("main"), /no student rows/);
  await shot(page, "11-import-error");

  await upload(page, "sample-marks.xlsx");
  await page.setInputFiles("#fileInput", { name: "notes.xlsx", mimeType: "application/octet-stream", buffer: Buffer.from("PK\x03\x04garbage-not-a-zip") });
  await page.waitForFunction(() => !window.__gc.state.load);
  assert.match(await page.textContent("[role=alert]"), /couldn’t be opened as a workbook[\s\S]*previous import \(sample-marks.xlsx\) is still loaded/);
  await page.setInputFiles("#fileInput", { name: "marks.pdf", mimeType: "application/pdf", buffer: Buffer.from("%PDF") });
  assert.match(await page.textContent("[role=alert]"), /isn’t an Excel workbook/);
  assert.equal(await page.evaluate(() => window.__gc.state.parsed.counts.valid), 248);
  assert.deepEqual(errors, []);
  await close();
});

test("second file replaces the course list; same file twice re-imports; ranges survive", async () => {
  const { page, errors, close } = await open();
  await upload(page, "sample-marks.xlsx");
  await page.click("text=Review data →");
  await page.click('.rail__item button:has-text("Operating Systems")');
  await page.click("text=Continue to grading →");
  await page.fill("#band-0-min", "78");

  await page.click("#fileChip button");
  await upload(page, "second-file.xlsx");
  await page.click("text=Review data →");
  const names = await page.$$eval(".rail__name", (e) => e.map((x) => x.textContent));
  assert.deepEqual(names, ["Computer Networks"]);
  assert.equal(await page.textContent("h1"), "Computer Networks");

  // Re-import the first file twice in a row; each import must register.
  await stepBtn(page, "Import").click();
  await upload(page, "sample-marks.xlsx");
  const t1 = await page.evaluate(() => window.__gc.state.importedAt.getTime());
  await page.waitForTimeout(20);
  await upload(page, "sample-marks.xlsx");
  const t2 = await page.evaluate(() => window.__gc.state.importedAt.getTime());
  assert.ok(t2 > t1, "same file imported twice");
  assert.match(await page.textContent("#toast"), /grade ranges for Operating Systems were kept/);
  const bands = await page.evaluate(() => window.__gc.state.bandsByCourse.get("Operating Systems")[0].min);
  assert.equal(bands, 78);
  assert.deepEqual(errors, []);
  await close();
});

test("CSV-hostile values are escaped and formulas neutralised", async () => {
  const { page, close } = await open();
  await upload(page, "csv-hostile.xlsx");
  await page.click("text=Review data →");
  await page.click("text=Continue to grading →");
  await stepBtn(page, "Export").click();
  await page.fill("#instructor", 'Rao, K. "Senior"');
  const [dl] = await Promise.all([page.waitForEvent("download"), page.click("#exportBtn")]);
  const csv = readFileSync(await dl.path(), "utf8").replace(/^﻿/, "");
  assert.equal(csv.split("\r\n")[0], 'Instructor,"Rao, K. ""Senior"""');
  assert.equal(csv.split("\r\n")[1], 'Course,"Course, With ""Comma"""');
  assert.match(csv, /^"'=HYPERLINK\(""http:\/\/x"",""click""\)",77,A-$/m);
  await close();
});

test("identical marks and a 20,000-row file render without errors", async () => {
  const { page, errors, close } = await open();
  await upload(page, "identical-marks.xlsx");
  await page.click("text=Review data →");
  assert.match(await page.textContent(".insights"), /Every student scored 75/);
  const t = Date.now();
  await stepBtn(page, "Import").click();
  await upload(page, "large-20k.xlsx");
  const ms = Date.now() - t;
  assert.match(await page.textContent(".result__line"), /20,000 student records.*40 courses/);
  await page.click("text=Review data →");
  await page.click("text=Continue to grading →");
  const svgH = await page.$eval("#bandChart svg", (s) => s.getBoundingClientRect().height);
  assert.ok(svgH < 300, "chart stays within its box");
  assert.ok(ms < 15000, `large import took ${ms}ms`);
  assert.deepEqual(errors, []);
  await close();
});

test("'Try with synthetic sample data' imports the bundled sample", async () => {
  const { page, errors, close } = await open();
  await page.click("text=Try with synthetic sample data");
  await page.waitForSelector(".result__line", { timeout: 30000 });
  assert.match(await page.textContent(".result__line"), /248 student records.*3 courses.*0 errors/);
  assert.deepEqual(errors, []);
  await close();
});

test("mobile width: no horizontal page scroll on any step", async () => {
  const { page, errors, close } = await open({ width: 375, height: 800 });
  const noScroll = () => page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth);
  assert.ok(await noScroll(), "import");
  await upload(page, "sample-marks.xlsx");
  await shot(page, "12-mobile-import");
  for (const s of ["Review", "Grade", "Export"]) {
    await stepBtn(page, s).click();
    assert.ok(await noScroll(), s);
    await shot(page, "13-mobile-" + s.toLowerCase());
  }
  assert.deepEqual(errors, []);
  await close();
});
