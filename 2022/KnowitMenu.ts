/**
 * Adds a custum menu to the document
 * Updated 2022-11-08
 * @author <thomas.malt@knowit.no>
 */
function onOpen() {
  const update = "2022-11-08";
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const entries = [
    {
      name: 'Update Competency Data',
      functionName: 'updateCompetencyData',
    },
    { name: 'Generate Data Sheet', functionName: 'generateDataSheet' },
  ];
  sheet.addMenu('knowit.no', entries);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('T-shape');
  if (sh === null) throw new TypeError('Sheet T-shape was null');

  sh.getRange(1, 1).setValue((new Date()).toJSON());
}
