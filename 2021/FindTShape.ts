/**
 * Calculates the primary, secondary and tertiary competency for a user
 *
 * @returns Array
 * @customfunction
 */
function calcTShape(data: any[][]) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('T-shape');
  if (sh === null) throw new TypeError('Sheet T-shape was null');

  const ra = sh.getRange(2, 3, 1, 9);
  const colors = ra.getBackgrounds();
  console.log('color', colors);
  const categories = ra.getValues().flat();
  const output = data.map((r) => {
    return r
      .map((e, i) => [e, i])
      .sort((a, b) => (a[0] < b[0] ? 1 : -1))
      .map((e) => categories[e[1]].substr(0, 12))
      .slice(0, 3);
  });

  console.log('output', output);
  return output;
}
