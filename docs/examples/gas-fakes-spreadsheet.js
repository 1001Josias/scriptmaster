function main() {
  const spreadsheet = SpreadsheetApp.openById('spreadsheet-id');
  const sheet = spreadsheet.getSheetByName('Data');

  if (!sheet) {
    throw new Error('Sheet not found: Data');
  }

  const values = sheet.getRange('A1:B2').getValues();
  Logger.log('Read %s rows', values.length);
  return values;
}
