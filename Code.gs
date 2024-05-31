let TRANSACTIONS_SHEET_ID = '85987263';

/**
 * @return {SpreadsheetApp.Sheet}
 */
function TransactionsSheet(){
  let sheets = SpreadsheetApp.getActive().getSheets();
  let txSheet = sheets.filter(s => s.getSheetId() == TRANSACTIONS_SHEET_ID)[0];
  if (!txSheet) throw new Error("Could not find transaction sheet! Check config");
  return txSheet;
}