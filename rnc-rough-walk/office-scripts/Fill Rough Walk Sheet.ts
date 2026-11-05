/**
 * Fill Rough Walk Sheet
 *
 * Run by the "RQI Walk Submitted" flow on a temporary copy of
 * RQI Rough Walk Template.xlsx. The copy is then converted to PDF.
 *
 * dataJson: a JSON object. Each key is a cell name in the workbook
 *   (Formulas > Name Manager), for example {"Builder":"ARCHER","Lot":"169"}.
 *   - A name that covers 3 cells in one row is a Yes / No / N/A answer:
 *     X goes in the first cell for Yes, the second for No, the third for N/A.
 *   - Any other name gets the value as text in its first cell.
 * signatureBase64: the signature PNG as base64 (no "data:" prefix). May be blank.
 *
 * Returns the names that were not found in the workbook, separated by commas
 * (blank when everything was filled).
 */
interface WalkData {
  [key: string]: string;
}

function main(workbook: ExcelScript.Workbook, dataJson: string, signatureBase64: string): string {
  let data: WalkData = JSON.parse(dataJson) as WalkData;
  let keys: string[] = Object.keys(data);
  let missing: string[] = [];

  for (let i = 0; i < keys.length; i++) {
    let key: string = keys[i];
    let value: string = data[key] === null || data[key] === undefined ? "" : String(data[key]);
    let item = workbook.getNamedItem(key);
    if (!item) {
      missing.push(key);
      continue;
    }
    let range = item.getRange();

    if (range.getRowCount() === 1 && range.getColumnCount() === 3) {
      fillChoice(range, value);
    } else {
      let cell = range.getCell(0, 0);
      cell.setNumberFormat("@");
      cell.setValue(value);
    }
  }

  if (signatureBase64 && signatureBase64.length > 0) {
    let sigItem = workbook.getNamedItem("Signature");
    if (!sigItem) {
      missing.push("Signature");
    } else {
      placeSignature(sigItem.getRange(), signatureBase64);
    }
  }

  return missing.join(", ");
}

/** Puts X under Yes, No or N/A and clears the other two cells. */
function fillChoice(range: ExcelScript.Range, value: string) {
  let answer: string = value.trim().toUpperCase();
  let col: number = -1;
  if (answer === "YES") {
    col = 0;
  } else if (answer === "NO") {
    col = 1;
  } else if (answer === "N/A" || answer === "NA") {
    col = 2;
  }
  for (let c = 0; c < 3; c++) {
    let cell = range.getCell(0, c);
    cell.setNumberFormat("@");
    cell.setValue(c === col ? "X" : "");
  }
}

/** Adds the signature image inside the Signature cells, scaled to fit and centred. */
function placeSignature(range: ExcelScript.Range, base64: string) {
  let sheet = range.getWorksheet();
  let boxLeft: number = range.getLeft();
  let boxTop: number = range.getTop();
  let boxWidth: number = range.getWidth();
  let boxHeight: number = range.getHeight();

  let shape = sheet.addImage(base64);
  shape.setLockAspectRatio(true);

  let imgWidth: number = shape.getWidth();
  let imgHeight: number = shape.getHeight();
  if (imgWidth <= 0 || imgHeight <= 0) {
    return;
  }

  let scale: number = Math.min(boxWidth / imgWidth, boxHeight / imgHeight);
  let newWidth: number = imgWidth * scale;
  let newHeight: number = imgHeight * scale;

  shape.setWidth(newWidth);
  shape.setHeight(newHeight);
  shape.setLeft(boxLeft + (boxWidth - newWidth) / 2);
  shape.setTop(boxTop + (boxHeight - newHeight) / 2);
}
