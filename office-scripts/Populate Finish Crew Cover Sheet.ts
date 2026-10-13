/**
 * Populate Finish Crew Cover Sheet (one row per job)
 *
 * Fills the MASTER COVER sheet of FINISH COVER SHEET MASTER.xlsx:
 *   - B4: job date (written as a real Excel date so the cell's date format is kept)
 *   - Table FINISHCOVER (header row 6): one row per submission.
 *
 * Columns are filled BY HEADER NAME, so the template's columns can be added to or reordered.
 * Recognised headers (case doesn't matter):
 *   Job Number | Supervisor | Lead Installer | Helper Installer | Builder | Subdivision |
 *   Lot # | Phase | Address | General Comments | Builder Comments | Field Super Comments |
 *   FQI Comments
 * Any other column is left blank.
 *
 * The template's lookup columns hold XLOOKUP formulas to the P: drive Production Schedule. They
 * are cleared and replaced with the values from the app, so the report never shows "Not found".
 *
 * Parameters (unchanged, so the flow needs no changes):
 *   recordsJson - JSON array of SUBMISSIONS items (SharePoint internal names)
 *   reportDate  - "yyyy-mm-dd" (or any date text Excel can read)
 */

interface SPChoice {
  Value?: string;
  value?: string;
}

interface SubmissionRecord {
  Title?: string;
  JobNumber?: string;
  Job_x0020_Number?: string;

  Supervisor?: string;

  LeadInstaller?: string;
  HelperInstaller?: string;

  BuilderName?: string;
  Builder?: SPChoice | string;

  SUBDIVISION?: string;
  Subdivision?: string;

  LotNumber?: string;

  Phase?: string;
  PHS?: string;

  // Street Address column (internal name JobAddress).
  JobAddress?: string;
  Address?: string;

  // General Comments is stored in the column whose internal name is SupervisorComments.
  GeneralComments?: string;
  SupervisorComments?: string;

  BuilderComments?: string;

  FieldSupervisorComments?: string;
  FieldSuperComments?: string;
  Field_x0020_Supervisor_x0020_Comments?: string;

  FQIComments?: string;
}

const SHEET_NAME = "MASTER COVER";
const TABLE_NAME = "FINISHCOVER";
const DATE_CELL = "B4";

function main(workbook: ExcelScript.Workbook, recordsJson: string, reportDate: string) {
  const ws = workbook.getWorksheet(SHEET_NAME);
  if (!ws) {
    throw new Error("Worksheet '" + SHEET_NAME + "' was not found.");
  }
  const table = workbook.getTable(TABLE_NAME);
  if (!table) {
    throw new Error("Table '" + TABLE_NAME + "' was not found.");
  }

  let records: SubmissionRecord[] = [];
  if (recordsJson && recordsJson.trim() !== "") {
    records = JSON.parse(recordsJson) as SubmissionRecord[];
  }

  // ---- Job date --------------------------------------------------------
  const serial = toExcelDate(reportDate);
  if (serial !== null) {
    ws.getRange(DATE_CELL).setValue(serial);
  } else {
    ws.getRange(DATE_CELL).setValue(reportDate || "");
  }

  // ---- Table layout ----------------------------------------------------
  const headerRange = table.getHeaderRowRange();
  const headerTexts = headerRange.getTexts()[0];
  const headers: string[] = [];
  for (let h = 0; h < headerTexts.length; h++) {
    headers.push(headerTexts[h].trim().toLowerCase());
  }
  const colCount = headers.length;
  const firstCol = headerRange.getColumnIndex();                 // 0-based
  const firstLetter = columnLetter(firstCol);
  const lastLetter = columnLetter(firstCol + colCount - 1);
  const headerRow = headerRange.getRowIndex() + 1;               // 1-based
  const firstDataRow = headerRow + 1;

  // ---- Build one row of values per record ------------------------------
  const rows: string[][] = [];
  for (let ri = 0; ri < records.length; ri++) {
    const values: string[] = [];
    for (let hi = 0; hi < headers.length; hi++) {
      values.push(valueFor(headers[hi], records[ri]));
    }
    rows.push(values);
  }

  // ---- Make sure the table is big enough -------------------------------
  const tableRange = table.getRange();
  let lastTableRow = tableRange.getRowIndex() + tableRange.getRowCount(); // 1-based last row
  const neededLastRow = headerRow + Math.max(rows.length, 1);
  if (neededLastRow > lastTableRow) {
    table.resize(ws.getRange(firstLetter + headerRow + ":" + lastLetter + neededLastRow));
    lastTableRow = neededLastRow;
  }

  // ---- Clear old values and the P: drive lookup formulas ----------------
  const body = ws.getRange(firstLetter + firstDataRow + ":" + lastLetter + lastTableRow);
  body.clear(ExcelScript.ClearApplyTo.contents);

  // Job Number and Lot # as text so values like "0110" or "9B" are kept exactly.
  const textCols = ["job number", "lot #", "lot number"];
  for (let i = 0; i < headers.length; i++) {
    if (textCols.indexOf(headers[i]) >= 0) {
      const letter = columnLetter(firstCol + i);
      ws.getRange(letter + firstDataRow + ":" + letter + lastTableRow).setNumberFormat("@");
    }
  }

  // Fill every table row: records first, then blanks. Writing the whole block replaces any
  // remaining calculated-column formulas with plain values.
  const allRows: string[][] = [];
  const bodyRowCount = lastTableRow - firstDataRow + 1;
  for (let i = 0; i < bodyRowCount; i++) {
    if (i < rows.length) {
      allRows.push(rows[i]);
    } else {
      allRows.push(emptyRow(colCount));
    }
  }
  body.setValues(allRows);

  // ---- Layout -----------------------------------------------------------
  const lastRow = rows.length > 0 ? firstDataRow + rows.length - 1 : firstDataRow;

  // Wrap long text so comments and wrapped names are fully visible, then fit row heights.
  const filled = ws.getRange(firstLetter + firstDataRow + ":" + lastLetter + lastRow);
  filled.getFormat().setWrapText(true);
  filled.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  filled.getFormat().autofitRows();

  // Print only the filled rows; repeat the column headings on every page.
  const layout = ws.getPageLayout();
  layout.setPrintArea("A1:" + lastLetter + lastRow);
  layout.setPrintTitleRows("$" + headerRow + ":$" + headerRow);
}

/** The value for one report column, chosen by its header text (lower case). */
function valueFor(header: string, r: SubmissionRecord): string {
  switch (header) {
    case "job number":
      return firstText([r.Title, r.JobNumber, r.Job_x0020_Number]);
    case "supervisor":
      return getText(r.Supervisor);
    case "lead installer":
      return getText(r.LeadInstaller);
    case "helper installer":
      return getText(r.HelperInstaller);
    case "builder":
      return firstText([r.BuilderName, r.Builder]);
    case "subdivision":
      return firstText([r.SUBDIVISION, r.Subdivision]);
    case "lot #":
    case "lot number":
      return getText(r.LotNumber);
    case "phase":
      return firstText([r.Phase, r.PHS]);
    case "address":
    case "street address":
      return firstText([r.JobAddress, r.Address]);
    case "general comments":
      return firstText([r.GeneralComments, r.SupervisorComments]);
    case "builder comments":
      return getText(r.BuilderComments);
    case "field super comments":
    case "field supervisor comments":
      return firstText([r.FieldSupervisorComments, r.FieldSuperComments, r.Field_x0020_Supervisor_x0020_Comments]);
    case "fqi comments":
      return getText(r.FQIComments);
    default:
      return "";
  }
}

/** 0-based column index to a column letter (0 = A, 25 = Z, 26 = AA). */
function columnLetter(index: number): string {
  let n = index + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/** A row of empty strings. */
function emptyRow(count: number): string[] {
  const row: string[] = [];
  for (let i = 0; i < count; i++) {
    row.push("");
  }
  return row;
}

/** Text from a string, number or SharePoint choice/lookup object ({ Value: "..." }). */
function getText(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "object") {
    const obj = value as { Value?: unknown; value?: unknown };
    if (obj.Value !== undefined && obj.Value !== null) {
      return String(obj.Value).trim();
    }
    if (obj.value !== undefined && obj.value !== null) {
      return String(obj.value).trim();
    }
    return "";
  }
  return String(value).trim();
}

/** First non-blank value in the list. */
function firstText(values: unknown[]): string {
  for (let i = 0; i < values.length; i++) {
    const t = getText(values[i]);
    if (t !== "") {
      return t;
    }
  }
  return "";
}

/** "yyyy-mm-dd", "yyyy-mm-ddThh:mm..." or "m/d/yyyy" to an Excel date serial; null if unreadable. */
function toExcelDate(input: string): number | null {
  if (!input) {
    return null;
  }
  const text = String(input).trim();
  let y = 0;
  let m = 0;
  let d = 0;

  const datePart = text.split("T")[0];
  const dash = datePart.split("-");
  const slash = datePart.split("/");
  if (dash.length === 3 && dash[0].length === 4) {
    y = Number(dash[0]);
    m = Number(dash[1]);
    d = Number(dash[2]);
  } else if (slash.length === 3) {
    m = Number(slash[0]);
    d = Number(slash[1]);
    y = Number(slash[2]);
    if (y < 100) {
      y += 2000;
    }
  } else {
    return null;
  }

  if (isNaN(y) || isNaN(m) || isNaN(d)) {
    return null;
  }
  const utc = Date.UTC(y, m - 1, d);
  const check = new Date(utc);
  if (check.getUTCFullYear() !== y || check.getUTCMonth() !== m - 1 || check.getUTCDate() !== d) {
    return null;
  }
  return utc / 86400000 + 25569;
}
