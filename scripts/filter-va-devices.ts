
/**
 * filter-va-devices.ts
 * 
 * Office Script — Acronis VA Device Report Filter
 * Author: Stylianos Vardakas
 * Version: 1.0
 * 
 * Description:
 * Filters the Acronis full device list Excel report, removing all rows
 * where the "Device name" column does not contain the identifier "-VA-"
 * in order to keep only Virtual Appliances.
 *
 * Designed to be executed via Power Automate immediately after the
 * report is saved to SharePoint.
 */
 
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getFirstWorksheet();
  const range = sheet.getUsedRange();
  const values = range.getValues();
  const headers = values[0] as string[];
  const colIndex = headers.findIndex(h => h === "Device name");
  if (colIndex === -1) throw new Error("Column 'Device name' not found.");
  for (let i = values.length - 1; i >= 1; i--) {
    const cellValue = String(values[i][colIndex]);
    if (!cellValue.includes("-VA-")) {
      sheet.getRange(`${i + 1}:${i + 1}`).delete(ExcelScript.DeleteShiftDirection.up);
    }
  }
}
