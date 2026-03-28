# PEN ID Excel Auto Updater (Chrome Extension)

Chrome extension that:

1. Reads all rows from the uploaded Excel file.
2. Finds each matching `PEN ID` row on the active student table webpage.
3. Double-clicks the row to open the edit screen.
4. Fills form fields from Excel values, clicks submit, returns to table, and repeats.
5. Can also export the detected student table on the current webpage to an Excel file.
6. Can export detailed Screening Details per PEN ID (Weight, Height, BMI, BMI Classification, etc.).

## Folder Structure

- `manifest.json` - extension manifest (MV3)
- `popup/popup.html` - popup UI
- `popup/popup.css` - popup styling
- `src/content.js` - webpage automation + scraper
- `src/popup.js` - Excel parser + orchestration logic
- `scripts/build.mjs` - build script for extension JS
- `sample-page.html` - sample page for testing

## Setup

1. Install dependencies:

```bash
npm install
```

2. Build extension scripts:

```bash
npm run build
```

3. Load in Chrome:
   - Open `chrome://extensions/`
   - Enable **Developer mode**
   - Click **Load unpacked**
   - Select this project folder

## Excel Requirements

Your Excel file must have one PEN column:

- `PEN ID` (recommended), or
- `PEN_ID`, or
- `pen_id`, or
- `PEN`

For best results, keep Excel column names aligned with the webpage edit form column names.

## How To Use

1. Open the target webpage.
2. Click the extension icon (opens a dedicated app window).
3. Choose your Excel file (`.xlsx`, `.xls`, or `.csv`).
4. Click **Upload & Auto Fill All Rows**.
5. Wait for completion summary (success/failed PEN IDs).
6. Click **Close** in the app window when finished.

## Export Table To Excel

1. Open the student table webpage in Chrome.
2. Click the extension icon.
3. Click **Download Table as Excel**.
4. The extension detects a table with columns like PEN ID / Name / SN and downloads it.

## Export Detailed Screening Excel

1. Open the student table webpage (where each row has a Screening Actions control).
2. Click the extension icon.
3. Click **Download Detailed Excel**.
4. The extension opens each student Screening Details view, scrapes detail values, computes BMI, returns to table, and continues.

## Notes

- Extension reads the first sheet in the workbook.
- It processes rows that contain valid PEN IDs.
- For each PEN ID, it automates: open edit -> fill -> submit -> back.
