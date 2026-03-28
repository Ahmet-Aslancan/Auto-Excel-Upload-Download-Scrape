import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx";

const ROOT = process.cwd();
const INPUT_FILE = path.join(ROOT, "Project sheet School.xlsx");
const OUTPUT_FILE = path.join(ROOT, "tables-from-excel.html");

const TARGET_COLUMNS = [
  "SN",
  "PEN ID",
  "Name",
  "Gender",
  "DOB",
  "Age",
  "Class",
  "Father's Name",
  "Mother's Name",
  "ABHA Number",
  "Mobile Number",
  "Screening Actions",
  "Weight (kg) *",
  "Height/Length (cm) *",
  "Blood Pressure",
  "Vision - Left Eye",
  "Vision - Right Eye",
  "Birth Defect Found",
  "Referal",
  "Select Defect type",
  "Defect Other",
  "Identification Code"
];

function normalize(value) {
  return String(value || "")
    .replace(/\u00A0/g, " ")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function isRowEmpty(row) {
  return !(row || []).some((cell) => String(cell || "").trim() !== "");
}

function detectHeaderRows(rows) {
  const indexes = [];
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const hasPenId = row.some((cell) => normalize(cell) === "penid" || normalize(cell) === "pen");
    const hasSn = row.some((cell) => normalize(cell) === "sn");
    if (hasPenId && hasSn) indexes.push(i);
  }
  return indexes;
}

function findTableTitle(rows, headerRowIndex, fallbackTitle) {
  for (let i = headerRowIndex - 1; i >= 0; i -= 1) {
    const row = rows[i] || [];
    const values = row.map((c) => String(c || "").trim()).filter(Boolean);
    if (!values.length) continue;
    const first = values[0];
    if (/institute\s*name/i.test(first) && values[1]) return values[1];
    if (!/sn|pen\s*id/i.test(first)) return first;
  }
  return fallbackTitle;
}

function extractTables(rows) {
  const headerIndexes = detectHeaderRows(rows);
  if (!headerIndexes.length) return [];

  return headerIndexes.map((headerRowIndex, idx) => {
    const nextHeader = headerIndexes[idx + 1] ?? rows.length;
    const originalHeaders = (rows[headerRowIndex] || []).map((h) => String(h || "").replace(/\u00A0/g, " ").trim());

    const mappedIndexes = TARGET_COLUMNS.map((target) =>
      originalHeaders.findIndex((source) => normalize(source) === normalize(target))
    );
    const penIdColumnIndex = TARGET_COLUMNS.findIndex((c) => normalize(c) === "penid");

    const bodyRows = [];
    for (let r = headerRowIndex + 1; r < nextHeader; r += 1) {
      const row = rows[r] || [];
      if (isRowEmpty(row)) continue;

      const mappedRow = TARGET_COLUMNS.map((_col, i) => {
        const srcIdx = mappedIndexes[i];
        return srcIdx >= 0 ? row[srcIdx] ?? "" : "";
      });

      const penIdValue = mappedRow[penIdColumnIndex];
      if (!String(penIdValue || "").trim()) continue;

      const hasAtLeastOneValue = mappedRow.some((v) => String(v || "").trim() !== "");
      if (hasAtLeastOneValue) bodyRows.push(mappedRow);
    }

    return {
      title: findTableTitle(rows, headerRowIndex, `Table ${idx + 1}`),
      headers: TARGET_COLUMNS,
      rows: bodyRows
    };
  });
}

function renderHtml(tables) {
  const sections = tables
    .map((table, tableIdx) => {
      const rows = table.rows
        .map(
          (row) =>
            `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`
        )
        .join("");

      return `
      <section class="table-card">
        <div class="table-header">
          <h2>${escapeHtml(table.title || `Table ${tableIdx + 1}`)}</h2>
          <span>${table.rows.length} record(s)</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>${table.headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr>
            </thead>
            <tbody>
              ${rows || `<tr><td colspan="${table.headers.length}" class="empty">No records</td></tr>`}
            </tbody>
          </table>
        </div>
      </section>`;
    })
    .join("\n");

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Excel Tables Preview</title>
    <style>
      :root {
        --bg: #f4f6ff;
        --card: #ffffff;
        --text: #1a1f36;
        --muted: #6b7280;
        --border: #dfe3f0;
        --brand: #5b34e6;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        font-family: Inter, Segoe UI, Arial, sans-serif;
        color: var(--text);
        background: linear-gradient(155deg, #eef2ff 0%, #f8f5ff 100%);
      }
      .container {
        max-width: 1280px;
        margin: 0 auto;
        padding: 24px 20px 50px;
      }
      .hero {
        background: linear-gradient(120deg, #5b34e6, #7f4dff 70%, #10b3c7);
        color: #fff;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 16px;
      }
      .hero h1 { margin: 0 0 6px; font-size: 24px; }
      .hero p { margin: 0; opacity: 0.95; }
      .meta {
        display: inline-block;
        margin-top: 10px;
        background: rgba(255, 255, 255, 0.18);
        border: 1px solid rgba(255, 255, 255, 0.35);
        border-radius: 999px;
        padding: 5px 10px;
        font-size: 12px;
      }
      .table-card {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 12px;
        margin-bottom: 12px;
      }
      .table-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 10px;
      }
      .table-header h2 {
        margin: 0;
        font-size: 16px;
      }
      .table-header span {
        font-size: 12px;
        color: var(--muted);
      }
      .table-wrap {
        overflow: auto;
        border: 1px solid var(--border);
        border-radius: 10px;
      }
      table {
        border-collapse: collapse;
        width: 100%;
        min-width: 1800px;
        background: #fff;
      }
      th, td {
        border-bottom: 1px solid var(--border);
        border-right: 1px solid var(--border);
        padding: 8px 10px;
        text-align: left;
        font-size: 12px;
        white-space: nowrap;
      }
      th {
        position: sticky;
        top: 0;
        background: #f5f7ff;
        z-index: 1;
        font-weight: 700;
      }
      td.empty {
        text-align: center;
        color: var(--muted);
      }
      th:first-child, td:first-child { position: sticky; left: 0; background: #fff; z-index: 2; }
      thead th:first-child { background: #f5f7ff; z-index: 3; }
    </style>
  </head>
  <body>
    <main class="container">
      <section class="hero">
        <h1>Excel Tables Web Preview</h1>
        <p>Auto-generated HTML tables from your workbook. One HTML table for each detected Excel table block.</p>
        <span class="meta">${tables.length} table(s) detected</span>
      </section>
      ${sections}
    </main>
  </body>
</html>`;
}

if (!fs.existsSync(INPUT_FILE)) {
  throw new Error(`Input file not found: ${INPUT_FILE}`);
}

const workbook = XLSX.readFile(INPUT_FILE);
const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

const tables = extractTables(rows);
if (!tables.length) {
  throw new Error("No Excel tables were detected. Ensure there is a header row with SN and PEN ID.");
}

const html = renderHtml(tables);
fs.writeFileSync(OUTPUT_FILE, html, "utf8");

console.log(`Generated ${OUTPUT_FILE} with ${tables.length} table(s).`);
