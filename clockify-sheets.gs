// ============================================
// CONFIGURATION — Update these values
// ============================================
const now = new Date();
const year = now.getFullYear();
const month = now.getMonth();

const CONFIG = {
  API_KEY: "",
  WORKSPACE_ID: "",
  USER_ID: "",
  SHEET_NAME: `${year}-${String(now.getMonth() + 1).padStart(2, '0')}`, // Sheet name has date in yyyy-MM format, e.g. 2026-11, change it to your preference
  PAGE_SIZE: 200,

  // Document fields
  NAME: "",
  POSITION: "",
  PROJECT: "", // Leave empty to use Clockify project, or set to override, e.g. "MarAdTech"
};

const BASE_URL = "https://api.clockify.me/api/v1";

// Table layout — starts at C5 (row 5, col 3)
const TABLE_START_ROW = 5;
const TABLE_START_COL = 3; // Column C
const NUM_COLS = 7;
const MIN_DATA_ROWS = 15; // Minimum rows to match template, expands if more days

const HEADERS = [
  "Date",
  "Specjalista",
  "Projekt",
  "Opis wykonanych Prac, w tym rezultatów Prac (np. Produkty)",
  "Liczba przepracowanych godzin (łącznie)",
  "Liczba przepracowanych godzin poświęconych na wykonywanie Prac w trakcie pozostawania w gotowości do wykonywania Prac",
  "Liczba godzin pozostawania w gotowości do wykonywania Prac",
];

// ============================================
// FETCH CURRENT MONTH
// ============================================
function fetchCurrentMonth() {
  // Build UTC strings explicitly to avoid timezone shifting
  const startISO = `${year}-${String(month + 1).padStart(2, "0")}-01T00:00:00Z`;

  // Last day of current month
  const lastDay = new Date(year, month + 1, 0).getDate();
  const endISO = `${year}-${String(month + 1).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}T23:59:59Z`;

  const label = Utilities.formatDate(
    new Date(year, month, 1),
    Session.getScriptTimeZone(),
    "MMMM yyyy",
  );

  importEntries(startISO, endISO, label);
}

// ============================================
// FETCH BY DATE RANGE
// ============================================
function fetchByDateRange() {
  const ui = SpreadsheetApp.getUi();

  const startResponse = ui.prompt("Enter start date (YYYY-MM-DD):");
  if (startResponse.getSelectedButton() !== ui.Button.OK) return;

  const endResponse = ui.prompt("Enter end date (YYYY-MM-DD):");
  if (endResponse.getSelectedButton() !== ui.Button.OK) return;

  const startText = startResponse.getResponseText().trim();
  const endText = endResponse.getResponseText().trim();

  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(startText) ||
    !/^\d{4}-\d{2}-\d{2}$/.test(endText)
  ) {
    ui.alert("Invalid date format. Use YYYY-MM-DD.");
    return;
  }

  const startISO = `${startText}T00:00:00Z`;
  const endISO = `${endText}T23:59:59Z`;

  importEntries(startISO, endISO, `${startText} to ${endText}`);
}

// ============================================
// SHARED IMPORT LOGIC
// ============================================
function importEntries(startISO, endISO, label) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME);
  const entries = getAllTimeEntries(startISO, endISO);

  if (entries.length === 0) {
    SpreadsheetApp.getUi().alert(`No time entries found for ${label}.`);
    return;
  }

  // Clear entire sheet
  sheet.clearContents();
  sheet.clearFormats();

  const projects = getProjects();
  const specialist = `${CONFIG.NAME} - ${CONFIG.POSITION}`;

  // Group entries by date
  const dayMap = {};
  entries.forEach((e) => {
    const start = new Date(e.timeInterval.start);
    const dateKey = Utilities.formatDate(
      start,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );

    if (!dayMap[dateKey]) {
      dayMap[dateKey] = { totalSeconds: 0, descriptions: [], projectName: "" };
    }

    const end = e.timeInterval.end ? new Date(e.timeInterval.end) : null;
    const durationSec = e.timeInterval.duration
      ? parseDuration(e.timeInterval.duration)
      : end
        ? (end - start) / 1000
        : 0;

    dayMap[dateKey].totalSeconds += durationSec;

    if (e.description) {
      dayMap[dateKey].descriptions.push(e.description);
    }

    if (!dayMap[dateKey].projectName && e.projectId && projects[e.projectId]) {
      dayMap[dateKey].projectName = projects[e.projectId];
    }
  });

  const sortedDates = Object.keys(dayMap).sort();

  // Build data rows — use whichever is greater: actual days or minimum rows
  const totalRows = Math.max(sortedDates.length, MIN_DATA_ROWS);
  const dataRows = [];
  for (let i = 0; i < totalRows; i++) {
    if (i < sortedDates.length) {
      const dateKey = sortedDates[i];
      const day = dayMap[dateKey];
      const hours = parseFloat((day.totalSeconds / 3600).toFixed(2));
      const projectName = CONFIG.PROJECT || day.projectName || "";
      const description = day.descriptions.join("; ");

      dataRows.push([
        dateKey,
        specialist,
        projectName,
        description,
        hours,
        "",
        "",
      ]);
    } else {
      // Empty row — still show specialist and project
      dataRows.push(["", specialist, CONFIG.PROJECT || "", "", "", "", ""]);
    }
  }

  // --- WRITE HEADERS (row 5) ---
  const headerRow = TABLE_START_ROW;
  const headerRange = sheet.getRange(headerRow, TABLE_START_COL, 1, NUM_COLS);
  headerRange.setValues([HEADERS]);
  headerRange.setFontWeight("bold");
  headerRange.setWrap(true);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");

  // --- WRITE DATA ROWS ---
  const dataStartRow = TABLE_START_ROW + 1;
  const dataRange = sheet.getRange(
    dataStartRow,
    TABLE_START_COL,
    totalRows,
    NUM_COLS,
  );
  dataRange.setValues(dataRows);
  dataRange.setVerticalAlignment("middle");

  // Center Date column (col C)
  sheet
    .getRange(dataStartRow, TABLE_START_COL, totalRows, 1)
    .setHorizontalAlignment("center");

  // Center hours columns (cols G and H)
  sheet
    .getRange(dataStartRow, TABLE_START_COL + 4, totalRows, 2)
    .setHorizontalAlignment("center");

  // --- WRITE "SUMA GODZIN" ROW ---
  const sumRow = dataStartRow + totalRows;

  // Merge columns C–F (cols 3–6) for "Suma godzin" label
  const sumLabelRange = sheet.getRange(sumRow, TABLE_START_COL, 1, 4);
  sumLabelRange.merge();
  sumLabelRange.setValue("Suma godzin");
  sumLabelRange.setFontWeight("bold");
  sumLabelRange.setHorizontalAlignment("center");
  sumLabelRange.setVerticalAlignment("middle");

  // Total hours in column G (col 7) = "Liczba przepracowanych godzin (łącznie)"
  const sumValueCell = sheet.getRange(sumRow, TABLE_START_COL + 4, 1);
  sumValueCell.setFormula(`=SUM(G${dataStartRow}:G${sumRow - 1})`);
  sumValueCell.setFontWeight("bold");
  sumValueCell.setHorizontalAlignment("center");
  sumValueCell.setVerticalAlignment("middle");

  // Empty cells for last two columns in sum row
  sheet.getRange(sumRow, TABLE_START_COL + 5).setValue("");
  sheet.getRange(sumRow, TABLE_START_COL + 6).setValue("");

  // --- BORDERS — thick outer, normal inner for entire table ---
  const fullTableRange = sheet.getRange(
    headerRow,
    TABLE_START_COL,
    totalRows + 2,
    NUM_COLS,
  );

  // All cells get normal border
  fullTableRange.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID,
  );

  // Thick outer border
  fullTableRange.setBorder(
    true,
    true,
    true,
    true,
    null,
    null,
    "#000000",
    SpreadsheetApp.BorderStyle.SOLID_THICK,
  );

  // --- COLUMN WIDTHS ---
  sheet.setColumnWidth(TABLE_START_COL, 100); // C — Data
  sheet.setColumnWidth(TABLE_START_COL + 1, 250); // D — Specjalista (wider)
  sheet.setColumnWidth(TABLE_START_COL + 2, 120); // E — Projekt
  sheet.setColumnWidth(TABLE_START_COL + 3, 180); // F — Opis
  sheet.setColumnWidth(TABLE_START_COL + 4, 140); // G — Liczba godzin łącznie (wider)
  sheet.setColumnWidth(TABLE_START_COL + 5, 160); // H — Liczba godzin poświęconych (wider)
  sheet.setColumnWidth(TABLE_START_COL + 6, 130); // I — Liczba godzin gotowości

  // Set header row height for wrapped text
  sheet.setRowHeight(headerRow, 80);

  SpreadsheetApp.getUi().alert(
    `Done! Imported ${sortedDates.length} days for ${label}.`,
  );
}

// ============================================
// API HELPERS
// ============================================
function getAllTimeEntries(startDate, endDate) {
  let allEntries = [];
  let page = 1;
  let hasMore = true;

  while (hasMore) {
    let url =
      `${BASE_URL}/workspaces/${CONFIG.WORKSPACE_ID}/user/${CONFIG.USER_ID}/time-entries` +
      `?page-size=${CONFIG.PAGE_SIZE}&page=${page}`;

    if (startDate) url += `&start=${startDate}`;
    if (endDate) url += `&end=${endDate}`;

    const response = callClockify(url);
    const entries = JSON.parse(response.getContentText());

    allEntries = allEntries.concat(entries);
    hasMore = entries.length === CONFIG.PAGE_SIZE;
    page++;
  }

  return allEntries;
}

function getProjects() {
  const url = `${BASE_URL}/workspaces/${CONFIG.WORKSPACE_ID}/projects?page-size=500`;
  const response = callClockify(url);
  const data = JSON.parse(response.getContentText());
  const map = {};
  data.forEach((p) => (map[p.id] = p.name));
  return map;
}

function callClockify(url) {
  const options = {
    method: "get",
    headers: { "X-Api-Key": CONFIG.API_KEY },
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(
      `Clockify API error (${response.getResponseCode()}): ${response.getContentText()}`,
    );
  }
  return response;
}

// ============================================
// UTILITY: Parse ISO 8601 duration (PT1H30M15S)
// ============================================
function parseDuration(iso) {
  const match = iso.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!match) return 0;
  return (
    parseInt(match[1] || 0) * 3600 +
    parseInt(match[2] || 0) * 60 +
    parseInt(match[3] || 0)
  );
}

// ============================================
// UTILITY: Get or create a sheet tab
// ============================================
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ============================================
// CUSTOM MENU
// ============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⏱ Clockify")
    .addItem("Fetch Current Month", "fetchCurrentMonth")
    .addItem("Fetch by Date Range", "fetchByDateRange")
    .addToUi();
}
