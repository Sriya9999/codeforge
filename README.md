# Grading Console

**Live:** https://grading-console.pages.dev

A prototype grading console, rebuilt from a buggy prototype as a debugging and product-design exercise.
It is **not** an official BITS Pilani grading system. Marks files are processed entirely in the browser
and are never uploaded.

## The workflow

| Step | Purpose | What the instructor sees |
|---|---|---|
| **01 Import** | Bring in the marks | Drag-and-drop or file picker, expected-format table, read/parse progress, then *"248 student records · 3 courses · 0 errors"* and a per-course table. Every file problem (wrong type, corrupt, password-protected, empty, missing columns) gets a plain-language message and a next step |
| **02 Review** | Trust the data before grading it | Students, mean, median, highest, lowest and standard deviation. A data-quality strip (valid rows, rows to check, duplicate IDs, missing or invalid marks). **Records requiring attention** listed with their Excel row numbers. Histogram with mean and median markers, and factual observations |
| **03 Grade** | Set the grade ranges | Per-mark chart with grade bands shaded, and editable From/To bounds with live counts. The status reads *"No gaps · No overlaps · 0–100 fully covered"* or names the exact problem. Also lists students close to a boundary, and offers reset with Undo and copying ranges from another course |
| **04 Export** | A consequential action, checked first | Ready-to-export checklist (students, bands, unassigned, overlaps, gaps, excluded records), instructor name, acknowledgement of excluded records, a file preview, and a receipt afterwards. If anything changes after an export, the console flags it |

Design choices worth noting:

- The course rail on the left makes grading several courses one continuous flow. After each export,
  a toast offers the next course.
- Re-importing a corrected file **keeps your grade ranges** for matching course names.
- Nothing is dropped silently. A row is either graded, counted once (an exact duplicate), or listed
  with its reason and excluded, and export requires acknowledging any exclusions.
- The ticking timer and double `confirm()` dialogs are gone. Time since import appears only in the export receipt.

## Input format

First sheet containing these three headers (order, case and apostrophe style don't matter):

| Student’s BITS ID | Course | Total Marks |
|---|---|---:|
| 2024A7PS0001P | Course A | 82 |

Marks must be whole numbers from 0 to 100. Leave out students who should receive NC.
The Import step has a blank template to download.

## Export format

The layout matches the original prototype, so downstream tools keep working. It is RFC 4180-quoted,
UTF-8 with BOM, and formula-safe:

```
Instructor,Prof. A. Sharma
Course,Data Structures & Algorithms

BITS ID,Total Marks,Grade
2024A7PS0001P,77,A-
```

The file is named `grades_<course>_<YYYY-MM-DD>.csv`.

## Project layout

```
index.html              page shell
src/core.js             all grading rules: parsing, validation, stats, bands, CSV (no DOM, unit-tested)
src/app.js              UI state and rendering
src/styles.css          design tokens and components
vendor/xlsx.full.min.js SheetJS 0.20.3, vendored (the original's unpinned CDN URL resolves to 0.18.5)
original/index.html     the prototype, verbatim, for reference and reproduction
fixtures/               generated .xlsx test files (happy path and every edge case in the brief)
tests/                  unit tests, browser E2E tests, and the reproduction harness for the original
deploy/build-worker.mjs bundles the site into a single Cloudflare Worker
BUG_LOG.md              28 documented bugs: reproduction, root cause, fix, verification
```

## Develop and test

```bash
npm install
npm run fixtures        # regenerate fixtures/*.xlsx
npm test                # 20 unit tests + 8 browser flows (headless Chromium)
npm run reproduce       # replay the bugs against original/index.html
npx http-server -p 8080 # then open http://localhost:8080
```

## Deploy

```bash
npm run deploy          # builds deploy/out/worker.js and deploys it to Cloudflare Pages
BASE_URL=https://grading-console.pages.dev/ node --test tests/e2e.test.mjs
```

The site is a static page, so any static host works (`index.html`, `src/`, `vendor/`, `fixtures/sample-marks.xlsx`).
It is deployed to Cloudflare Pages as a single `_worker.js` with the files inlined, plus a strict Content-Security-Policy
(`script-src 'self'`).
