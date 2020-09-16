/**
 * Filters through competency rows matching criteria and returns the people
 * with motivation above criteria.
 *
 * @example SYSTEMUTVIKLING4SALES(3, 3)
 * @param {Number} competency
 * @param {Number} motivation
 * @return Array rows of people with competency above threshold and motivation above threshold
 * @customfunction
 */
function SYSTEMUTVIKLING4SALES(competency = 3, motivation = 3) {
  const sCom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kompetanse');
  const sMot = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Motivation');
  const sKat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kategorier');

  const categories = sKat.getRange(2, 1, sKat.getLastRow() - 1, 1).getValues();

  sCom
    .getDataRange()
    .getValues()
    .map((r) => {
      let item = [r[0]];
      r.forEach((i) => {});
      return item;
    });

  return categories;
}
