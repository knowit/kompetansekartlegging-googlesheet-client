/**
 * Calculates the primary, secondary and tertiary competency area for a user
 *
 * @returns Array
 * @customfunction
 */
function calcTShape(data: any[][]) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('T-shape');

  if (sh === null) throw new TypeError('Sheet T-shape was null');

  const ra = sh.getRange(2, 4, 1, 9);
  const categories = ra.getValues().flat();

  console.log('categories', categories);

  const output = data.map((r) => {
    return r
      .map((e, i) => [e, i])
      .filter((e) => e[0] >= 1.0)
      .sort((a, b) => b[0] - a[0])
      .map((e) => categories[e[1]].substr(0, 12))
      .slice(0, 3);
  });

  console.log('output', output);
  return output;
}
