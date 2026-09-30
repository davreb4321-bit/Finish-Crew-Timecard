/**
 * Populate Finish Crew Cover Sheet (one row per job)
 *
 * Fills the MASTER COVER sheet of FINISH COVER SHEET MASTER.xlsx:
 *   - B4: job date (written as a real Excel date so the cell's date format is kept)
 *   - Table FINISHCOVER (header row 6): one row per submission, columns
 *     Job Number | Supervisor | Lead Installer | Helper Installer | Builder | Subdivision |
 *     Lot # | Phase | General Comments | Builder Comments | Field Super Comments | FQI Comments
 *
 * The template's Supervisor / Builder / Subdivision / Lot # columns hold XLOOKUP formulas to the
 * P: drive Production Schedule. They are cleared and replaced with the values from the app, so the
 * report never shows "Not found".
 *
 * Parameters (unchanged from the previous script, so the flow needs no changes):
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
const HEADER_ROW = 6;          // table header row
const FIRST_DATA_ROW = 7;
const COLUMN_COUNT = 12;       // A:L
const DATE_CELL = "B4";

function main(workbook: ExcelScript.Workbook, recordsJson: string, reportDate: string) {
  const ws = workbook.getWorksheet(SHEET_NAME);
  if (!ws) {
    throw new Error("Worksheet '" + SHEET_NAME + "' was not found.");
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

  // ---- Build one row of values per record ------------------------------
  const rows: string[][] = records.map(r => [
    firstText([r.Title, r.JobNumber, r.Job_x0020_Number]),
    getText(r.Supervisor),
    getText(r.LeadInstaller),
    getText(r.HelperInstaller),
    firstText([r.BuilderName, r.Builder]),
    firstText([r.SUBDIVISION, r.Subdivision]),
    getText(r.LotNumber),
    firstText([r.Phase, r.PHS]),
    firstText([r.GeneralComments, r.SupervisorComments]),
    getText(r.BuilderComments),
    firstText([r.FieldSupervisorComments, r.FieldSuperComments, r.Field_x0020_Supervisor_x0020_Comments]),
    getText(r.FQIComments)
  ]);

  // ---- Make sure the table is big enough -------------------------------
  const table = workbook.getTable(TABLE_NAME);
  let lastTableRow = 100;
  if (table) {
    const tableRange = table.getRange();
    lastTableRow = tableRange.getRowIndex() + tableRange.getRowCount(); // 1-based last row
    const neededLastRow = HEADER_ROW + Math.max(rows.length, 1);
    if (neededLastRow > lastTableRow) {
      table.resize(ws.getRange("A" + HEADER_ROW + ":L" + neededLastRow));
      lastTableRow = neededLastRow;
    }
  }

  // ---- Clear old values and the P: drive lookup formulas ----------------
  const body = ws.getRange("A" + FIRST_DATA_ROW + ":L" + lastTableRow);
  body.clear(ExcelScript.ClearApplyTo.contents);

  // Job Number and Lot # as text so values like "0110" or "9B" are kept exactly.
  ws.getRange("A" + FIRST_DATA_ROW + ":A" + lastTableRow).setNumberFormat("@");
  ws.getRange("G" + FIRST_DATA_ROW + ":G" + lastTableRow).setNumberFormat("@");

  // Fill every table row: records first, then blanks. Writing the whole block replaces any
  // remaining calculated-column formulas with plain values.
  const allRows: string[][] = [];
  const bodyRowCount = lastTableRow - FIRST_DATA_ROW + 1;
  for (let i = 0; i < bodyRowCount; i++) {
    if (i < rows.length) {
      allRows.push(rows[i]);
    } else {
      allRows.push(emptyRow());
    }
  }
  body.setValues(allRows);

  // ---- Layout -----------------------------------------------------------
  const lastRow = rows.length > 0 ? FIRST_DATA_ROW + rows.length - 1 : FIRST_DATA_ROW;

  // Wrap long text so comments and wrapped names are fully visible, then fit row heights.
  const filled = ws.getRange("A" + FIRST_DATA_ROW + ":L" + lastRow);
  filled.getFormat().setWrapText(true);
  filled.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  filled.getFormat().autofitRows();

  // Print only the filled rows; repeat the column headings on every page.
  const layout = ws.getPageLayout();
  layout.setPrintArea("A1:L" + lastRow);
  layout.setPrintTitleRows("$" + HEADER_ROW + ":$" + HEADER_ROW);
}

/** A row of 12 empty strings. */
function emptyRow(): string[] {
  const row: string[] = [];
  for (let i = 0; i < COLUMN_COUNT; i++) {
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
