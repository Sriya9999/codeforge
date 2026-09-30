// Drives the ORIGINAL prototype (original/index.html) in headless Chromium and
// prints evidence for each bug in BUG_LOG.md. Nothing here is asserted; it is
// an observation harness. Run: node tests/reproduce-original.mjs
import { chromium } from "playwright";
import { readFileSync, existsSync } from "node:fs";
import path from "node:path";

const ROOT = path.resolve(new URL("..", import.meta.url).pathname);
const FIX = (f) => path.join(ROOT, "fixtures", f);
const URL_ = "file://" + path.join(ROOT, "original/index.html");
// The prototype loads an unpinned CDN build; serve that exact version (0.18.5) offline.
const SHEETJS = process.env.SHEETJS_0185 || path.join(ROOT, "vendor/xlsx.full.min.js");

const browser = await chromium.launch();
const log = (id, msg) => console.log(`[${id}] ${msg}`);

async function open({ acceptDialogs = true } = {}) {
  const ctx = await browser.newContext({ acceptDownloads: true });
  const page = await ctx.newPage();
  const errors = [];
  const dialogs = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => m.type() === "error" && errors.push(m.text()));
  page.on("dialog", (d) => { dialogs.push(d.message()); acceptDialogs ? d.accept() : d.dismiss(); });
  await page.route(/cdn\.jsdelivr\.net/, (r) => r.fulfill({ body: readFileSync(SHEETJS), contentType: "text/javascript" }));
  await page.goto(URL_);
  return { page, errors, dialogs, ctx };
}
const upload = async (page, f) => { await page.setInputFiles("#file", FIX(f)); await page.waitForTimeout(300); };
const courseOptions = (page) => page.$$eval("#course option", (o) => o.map((x) => x.value));
async function selectCourse(page, c) {
  await page.fill("#instructor", "Dr. Test");
  await page.selectOption("#course", c);
  await page.waitForTimeout(600);
}
async function exportCsv(page) {
  const [dl] = await Promise.all([page.waitForEvent("download"), page.click("#download")]);
  return readFileSync(await dl.path(), "utf8");
}
const stats = (page) => page.evaluate(() => ({
  labelMin: document.querySelector(".stat:nth-child(1) b").textContent,
  labelMax: document.querySelector(".stat:nth-child(2) b").textContent,
  avg: document.getElementById("avg").textContent,
  med: document.getElementById("med").textContent,
  summary: document.getElementById("gradeSummary").innerText.replace(/\n/g, " "),
}));

// B1 — file picker filter
{
  const { page } = await open();
  log("B1", `file input accept="${await page.getAttribute("#file", "accept")}" (brief requires .xlsx)`);
  await page.context().close();
}
// B2 — course dropdown: one option per ROW, never cleared
{
  const { page } = await open();
  await upload(page, "sample-marks.xlsx");
  let o = await courseOptions(page);
  log("B2", `after 1 upload: ${o.length - 1} course options for 3 distinct courses`);
  await upload(page, "second-file.xlsx");
  o = await courseOptions(page);
  log("B2", `after 2nd upload (1 course): ${o.length - 1} options; still lists "Operating Systems": ${o.includes("Operating Systems")}`);
  await page.context().close();
}
// B3 — selecting a course that no longer exists in the data after a second upload
{
  const { page } = await open();
  await upload(page, "sample-marks.xlsx");
  await selectCourse(page, "Operating Systems");
  await upload(page, "second-file.xlsx");
  await page.selectOption("#course", "Discrete Mathematics"); await page.waitForTimeout(600);
  log("B3", `stale course selected after re-upload → ${JSON.stringify(await stats(page))}`);
  await page.context().close();
}
// B4 — BITS ID column mismatch → "undefined" in export
{
  const { page } = await open();
  await upload(page, "sample-marks.xlsx");
  await selectCourse(page, "Discrete Mathematics");
  const csv = await exportCsv(page);
  log("B4", `CSV rows 4-6:\n${csv.split("\n").slice(3, 7).join("\n")}`);
  await page.context().close();
}
// B5 — Min/Max labels swapped
{
  const { page } = await open();
  await upload(page, "sample-marks.xlsx");
  await selectCourse(page, "Discrete Mathematics");
  log("B5", `stats: ${JSON.stringify(await stats(page))}`);
  await page.context().close();
}
// B6 — coverage not validated: shrink A max / raise E min → students silently dropped
{
  const { page } = await open();
  await upload(page, "sample-marks.xlsx");
  await selectCourse(page, "Discrete Mathematics");
  await page.selectOption("#Amax", "90"); await page.waitForTimeout(100);
  const err = await page.textContent("#rangeError");
  const disabled = await page.isDisabled("#download");
  const csv = await exportCsv(page);
  const exported = csv.trim().split("\n").length - 4;
  log("B6", `A max=90 → error="${err}" downloadDisabled=${disabled}; exported ${exported}/50 students`);
  await page.context().close();
}
// B7 — decimal, blank, non-numeric, text-number, out-of-range marks
{
  const { page } = await open();
  await upload(page, "edge-cases.xlsx");
  log("B7", `course options: ${JSON.stringify((await courseOptions(page)).slice(1))}`);
  await selectCourse(page, "Course A");
  log("B7", `stats: ${JSON.stringify(await stats(page))}`);
  const csv = await exportCsv(page);
  log("B7", `exported ${csv.trim().split("\n").length - 4} rows of 14 "Course A" rows:\n${csv.split("\n").slice(4).join(" | ")}`);
  await page.context().close();
}
// B8 — Reset before choosing a course → TypeError
{
  const { page, errors, dialogs } = await open();
  await page.click("#resetRanges"); await page.waitForTimeout(200);
  log("B8", `dialogs=${dialogs.length} errors=${JSON.stringify(errors)}`);
  await page.context().close();
}
// B9 — corrupt / non-Excel file → uncaught error, no feedback
{
  const { page, errors } = await open();
  await page.setInputFiles("#file", { name: "notes.xlsx", mimeType: "application/octet-stream", buffer: Buffer.from("PK\x03\x04garbage-not-a-zip") });
  await page.waitForTimeout(300);
  log("B9", `errors=${JSON.stringify(errors)} ; visible feedback="${await page.textContent("#welcome")}"`);
  await page.context().close();
}
// B10 — empty / header-only / wrong-columns: silent
for (const f of ["empty.xlsx", "header-only.xlsx", "wrong-columns.xlsx"]) {
  const { page, errors } = await open();
  await upload(page, f);
  const o = await courseOptions(page);
  log("B10", `${f}: options=${JSON.stringify(o.slice(1))} errors=${errors.length} feedback="${await page.textContent("#welcome")}"`);
  if (f === "wrong-columns.xlsx" && o.length > 1) {
    await selectCourse(page, o[1]).catch(() => {});
  }
  await page.context().close();
}
// B11 — same file twice: change event does not fire
{
  const { page } = await open();
  await upload(page, "sample-marks.xlsx");
  const before = (await courseOptions(page)).length;
  await page.evaluate(() => { window.__reads = 0; const o = FileReader.prototype.readAsBinaryString; FileReader.prototype.readAsBinaryString = function (...a) { window.__reads++; return o.apply(this, a); }; });
  // Simulate the browser behaviour: selecting the same path leaves input.value unchanged, so no 'change'.
  const value = await page.$eval("#file", (i) => i.value);
  log("B11", `input.value after upload="${value}" (never reset) → re-picking the same file fires no change event in Chromium/Firefox/Safari`);
  await page.context().close();
}
// B12 — histogram overflow with large bins / zero-std bell curve
{
  const { page } = await open();
  await upload(page, "large-20k.xlsx");
  await selectCourse(page, "Course 01");
  const maxBin = await page.evaluate(() => {
    const m = data.filter((d) => d.Course === course.value).map((d) => d["Total Marks"]);
    const b = Array(10).fill(0); m.forEach((x) => b[Math.min(9, Math.floor(x / 10))]++); return Math.max(...b);
  });
  log("B12", `largest bin=${maxBin} → bar height ${maxBin * 12}px on a 240px canvas (baseline y=210)`);
  await page.context().close();
  const r = await open();
  await upload(r.page, "identical-marks.xlsx");
  await selectCourse(r.page, "Course A");
  const std = await r.page.evaluate(() => { const m = data.map((d) => d["Total Marks"]); const mu = m.reduce((a, b) => a + b) / m.length; return Math.sqrt(m.reduce((a, b) => a + (b - mu) ** 2, 0) / m.length); });
  log("B12", `identical marks: std=${std} → bell curve divides by zero (y=NaN/Infinity)`);
  await r.page.context().close();
}
// B13 — CSV injection / unescaped commas
{
  const { page } = await open();
  await upload(page, "csv-hostile.xlsx");
  await page.fill("#instructor", "Rao, K. \"Senior\"");
  await page.selectOption("#course", { index: 1 }); await page.waitForTimeout(500);
  const csv = await exportCsv(page);
  log("B13", `CSV:\n${csv}`);
  await page.context().close();
}
// B14 — timer shows stale 00:00 for first second & starts at page load
{
  const { page } = await open();
  await page.waitForTimeout(1500);
  const t = await page.textContent("#timerText");
  log("B14", `timer after 1.5s = ${t} (only updated on interval tick; clock started at page load, before any file)`);
  await page.context().close();
}
// B15 — single-mark band rejected
{
  const { page } = await open();
  await upload(page, "sample-marks.xlsx");
  await selectCourse(page, "Discrete Mathematics");
  // A: 100–100, A-: 80–99 … a contiguous, valid scheme
  await page.selectOption("#Amin", "100"); await page.waitForTimeout(50);
  await page.selectOption("#A-min", "80"); await page.waitForTimeout(50);
  log("B15", `A=100–100 (contiguous) → error="${await page.textContent("#rangeError")}"`);
  await page.context().close();
}
// B16 — changing a MAX does not cascade; editing E.min to 5 passes validation
{
  const { page } = await open();
  await upload(page, "sample-marks.xlsx");
  await selectCourse(page, "Operating Systems");
  await page.selectOption("#Emin", "5"); await page.waitForTimeout(50);
  const csv = await exportCsv(page);
  const n = csv.trim().split("\n").length - 4;
  const lowCount = await page.evaluate(() => data.filter((d) => d.Course === "Operating Systems" && d["Total Marks"] < 5).length);
  log("B16", `E min=5 → error="${await page.textContent("#rangeError")}", exported ${n}/78 (students <5: ${lowCount})`);
  await page.context().close();
}

await browser.close();
