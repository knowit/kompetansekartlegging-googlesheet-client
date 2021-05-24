/**
 * Adds a custum menu to the document
 */
function onOpen() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const entries = [
        {
            name: "Update Competency Data", functionName: "updateCompetencyData",
        }
    ];
    sheet.addMenu("knowit.no", entries);
}