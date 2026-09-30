// Generates the .xlsx fixtures used for manual and automated testing.
// Run: node tests/make-fixtures.mjs
import * as XLSX from "xlsx";
import * as fs from "node:fs";
XLSX.set_fs(fs);
const { mkdirSync } = fs;

const OUT = new URL("../fixtures/", import.meta.url).pathname;
mkdirSync(OUT, { recursive: true });

const ID = "Student’s BITS ID"; // curly apostrophe, exactly as in the brief

// Deterministic PRNG so fixtures are reproducible.
let seed = 42;
const rand = () => ((seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296);
const normal = (mu, sd) => {
  const u = 1 - rand(), v = rand();
  return mu + sd * Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
};
const clamp = (x) => Math.max(0, Math.min(100, Math.round(x)));

function write(name, rows, header = [ID, "Course", "Total Marks"]) {
  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Marks");
  XLSX.writeFile(wb, OUT + name);
  console.log(`${name.padEnd(28)} ${rows.length} rows`);
}

function cohort(course, n, mu, sd, startId) {
  return Array.from({ length: n }, (_, i) => [
    `2024A7PS${String(startId + i).padStart(4, "0")}P`,
    course,
    clamp(normal(mu, sd)),
  ]);
}

// 1. Happy path: 248 students across 3 courses.
write("sample-marks.xlsx", [
  ...cohort("Data Structures & Algorithms", 120, 66, 15, 1),
  ...cohort("Operating Systems", 78, 58, 17, 201),
  ...cohort("Discrete Mathematics", 50, 71, 12, 301),
]);

// 2. A second, different file (tests re-upload / stale course list).
write("second-file.xlsx", cohort("Computer Networks", 64, 62, 14, 501));

// 3. Every data-quality problem from the brief, in one file.
write("edge-cases.xlsx", [
  ["2024A7PS0001P", "Course A", 82],
  ["2024A7PS0002P", "Course A", 71],
  ["2024A7PS0002P", "Course A", 64],     // duplicate ID, different mark
  ["2024A7PS0003P", "Course A", 55],
  ["2024A7PS0003P", "Course A", 55],     // exact duplicate record
  ["2024A7PS0004P", "Course A", -5],     // below 0
  ["2024A7PS0005P", "Course A", 104],    // above 100
  ["2024A7PS0006P", "Course A", 79.5],   // decimal (falls between 79 and 80)
  ["2024A7PS0007P", "Course A", null],   // blank mark
  ["2024A7PS0008P", "Course A", "AB"],   // non-numeric
  ["2024A7PS0009P", "Course A", "68"],   // number stored as text
  ["2024A7PS0010P", "Course A", "72"],   // number stored as text
  ["2024A7PS0011P", null, 60],           // missing course
  [null, "Course A", 45],                // missing ID
  ["2024A7PS0012P", "Course A ", 90],    // trailing-space course name
  ["2024A7PS0013P", "Course B", 40],
]);

// 4. Header row only.
write("header-only.xlsx", []);

// 5. Completely empty workbook.
{
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), "Sheet1");
  XLSX.writeFile(wb, OUT + "empty.xlsx");
  console.log("empty.xlsx                   0 rows");
}

// 6. Wrong / missing column names.
write("wrong-columns.xlsx", [["2024A7PS0001P", "Course A", 80]], ["ID No", "Subject", "Marks"]);
write("missing-marks-column.xlsx", [["2024A7PS0001P", "Course A"]], [ID, "Course"]);

// 7. Straight-apostrophe and plain "BITS ID" header variants (should be accepted).
write("header-variants.xlsx", [["2024A7PS0001P", "Course A", 88], ["2024A7PS0002P", "Course A", 34]],
  ["Student's BITS ID", " course ", "TOTAL MARKS"]);

// 8. Everyone has the same mark (zero standard deviation).
write("identical-marks.xlsx", Array.from({ length: 12 }, (_, i) =>
  [`2024A7PS${String(i + 1).padStart(4, "0")}P`, "Course A", 75]));

// 9. Large dataset: 20,000 students across 40 courses.
{
  const rows = [];
  for (let c = 0; c < 40; c++) rows.push(...cohort(`Course ${String(c + 1).padStart(2, "0")}`, 500, 50 + (c % 20), 16, c * 1000));
  write("large-20k.xlsx", rows);
}

// 10. CSV-hostile content (commas, quotes, formula prefixes).
write("csv-hostile.xlsx", [
  ["=HYPERLINK(\"http://x\",\"click\")", "Course, With \"Comma\"", 77],
  ["2024A7PS0002P", "Course, With \"Comma\"", 33],
]);
