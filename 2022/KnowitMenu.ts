/**
 * Adds a custum menu to the document
 */
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const entries = [
    {
      name: 'Update Competency Data',
      functionName: 'updateCompetencyData',
    },
    { name: 'Generate Data Sheet', functionName: 'generateDataSheet' },
  ];
  sheet.addMenu('knowit.no', entries);
}
