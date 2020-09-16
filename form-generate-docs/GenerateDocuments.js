/**
 * Knowit Objectnet Kompetansekartlegging - 2020 and beyond
 *
 * This Google Apps Script is part of the system for competency mapping at
 * Knowit Objectnet
 *
 * @author Thomas Malt <thomas.malt@knowit.no>
 * @copyright 2020 Knowit Objectnet AS
 */
const spreadsheets = {
  grupper: SpreadsheetApp.openByUrl(
    'https://docs.google.com/spreadsheets/d/1lHKZeE6UvGBYKHUpMEcHb4zMUkwyystWkCa1ED-F8yQ/edit'
  ),
};

const config = {
  version: '1.0.19',
  sheets: {
    svar: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Skjemasvar 1'),
    kategorier: spreadsheets.grupper.getSheetByName('Kategorier'),
    ansatte: spreadsheets.grupper.getSheetByName('Ansatt med leder'),
    ledere: spreadsheets.grupper.getSheetByName('Ledere'),
    radar: spreadsheets.grupper.getSheetByName('Radar'),
  },
  colors: {
    competencies: {
      header: '#c0e4e6', // "#46bdc6"
      gradient: {
        min: '#c5f3f7',
        max: '#46bdc6',
      },
    },
    motivation: {
      header: '#f8ebc4', // "#fbbc04",
      gradient: {
        min: '#f8ebc4',
        max: '#fbbc04',
      },
    },
    all: {
      gradient: {
        min: '#c0e4e6',
        max: '#04a4b0',
      },
    },
  },
};

/**
 * Event handler. Listens to change events and updates generated documents
 * when that happens
 */
function handleNewData(e) {
  console.log('Got new data trigger');

  updateDataMaster();
  const eid = findGoogleidInRow(e.range.rowStart);
  const mid = getManagerForEmployee(eid);

  // Regenerates all the documents (takes 5-10 minutes)
  // updateGroupManagerDocuments();

  // Only regenerate the group with changes
  updateGroupManagerDocumentForGoogleId(mid);
  console.log('Done.');
}

/**
 * Fetches the googleid for the new or changed row.
 *
 * @param {Number} row The row the change is in.
 * @returns {string} the google id email for the user.
 */
function findGoogleidInRow(row) {
  const googleid = config.sheets.svar.getRange(row, 2).getValue();
  console.log('find googleid in row:', googleid);
  return googleid;
}

/**
 * Does a lookup in the employee->manager spreadsheet and returns the googleid
 * of the first row that matches the employee id. (there should only be one).
 *
 * @param {string} id google id of employee
 * @returns {string} google id of manager
 * @customfunction
 */
function getManagerForEmployee(id) {
  const boss = config.sheets.ansatte
    .getDataRange()
    .getValues()
    .slice(3)
    .find((r) => r[2] === id)[6];

  console.log('get manager for employee:', id, ' -> ', boss);
  return boss;
}

/**
 * Iterates over all the group managers documents and updates them.
 * This function is nice to have to run manually if you want to regenerate all
 * documents for some reason.
 */
function updateGroupManagerDocuments() {
  console.log('Generating Group Manager Documents, version: ', config.version);
  getGroupManagers().forEach(updateGroupDocument);
}

/**
 * Regenerates the group manager document for a given google id.
 * If the function is called without arguments (manually from google scripts
 * for debug purposes a default group manager is used).
 *
 * @param {string} id google id of group manager
 * @returns string - the name of the generated file
 * @customfunction
 */
function updateGroupManagerDocumentForGoogleId(id = 'mayn.kjar@knowit.no') {
  const managers = getGroupManagers();
  return updateGroupDocument(managers.find((m) => m.googleid === id));
}

/**
 * Iterates over the spreadsheet with form responses and generates a master
 * data document to share with everyone
 */
function updateDataMaster() {
  const filename = `Kompetansekartlegging - 2020-02 - Alle Data - Anonymisert`;
  const ss = createOrFetchSpreadsheet(filename);
  const svar = config.sheets.svar;
  const rows = svar.getLastRow() - 1;
  const cols = svar.getLastColumn() - 1;

  const headers = svar.getRange(1, 2, 1, cols).getValues()[0];
  headers[0] = 'UUID';

  const data = svar.getRange(2, 2, rows, cols).getValues();
  const outSvar = ss.getSheetByName('Alle Svar') || ss.insertSheet('Alle Svar');

  outSvar.clear();
  outSvar
    .getRange(1, 1, 1, cols)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#efefef')
    .setFontSize(10)
    .setVerticalAlignment('bottom');
  outSvar
    .getRange(1, 2, 1, cols - 1)
    .setTextRotation(45)
    .setHorizontalAlignment('center');

  const outData = data
    .map((row) => {
      const email = row[0];
      const uuid = getHmacSHA256SigAsUUID(email);

      // thomas.malt@knowit.no 8c226907-956b-9f9e-921a-5b05ca475824
      row[0] = uuid;
      return row;
    })
    .sort();

  addDataMasterInformation(ss, outData);

  const dataRange = outSvar
    .getRange(2, 1, rows, cols)
    .setValues(outData)
    .setFontSize(8)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setFontFamily('Roboto Mono');
  outSvar.getRange(2, 1, rows, 1).setBackground('#efefef').setHorizontalAlignment('left');

  outSvar.setColumnWidths(2, cols - 1, 30);
  outSvar.setColumnWidth(1, 240);

  let dataRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(config.colors.all.gradient.max)
    .setGradientMinpoint(config.colors.all.gradient.min)
    .setRanges([dataRange])
    .build();

  let rules = outSvar.getConditionalFormatRules();
  rules.push(dataRule);
  outSvar.setConditionalFormatRules(rules);
}

/**
 * Adds a sheet with metadata.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet document
 * @param {Array<any>} data from the response spreadsheet.
 */
function addDataMasterInformation(ss, data) {
  const sheet = ss.getSheetByName('Informasjon') || ss.insertSheet('Informasjon');
  sheet.clear();

  sheet.getRange(1, 1, 3, 2).setValues([
    ['Dokument oppdatert', new Date().toJSON()],
    ['Script versjon', `v${config.version}`],
    ['Antall Svar', data.length],
  ]);

  sheet.setColumnWidths(1, 2, 200);
  sheet.getRange(1, 1, sheet.getLastRow(), 1).setFontWeight('bold');

  removeUnusedRows(sheet);
  removeUnusedColumns(sheet);
}

/**
 * Create or update document for group
 *
 * @param {any} gruppeleder
 */
function updateGroupDocument(gruppeleder) {
  // console.log("update group document for:", gruppeleder);
  const filename = `Kompetansekartlegging - 2020-02 - Gruppeoversikt - ${gruppeleder.visningsnavn}`;
  const ss = createOrFetchSpreadsheet(filename);

  ss.addEditor(gruppeleder.googleid);

  addGruppeoversikt(ss, gruppeleder);
  addKompetanseVsMotivasjon(ss, gruppeleder);
  // addTestGraph(ss, gruppeleder);
  addInformation(ss, gruppeleder);

  removeSheet(ss, 'Sheet1');

  return filename;
}

/**
 * Reorders the dataset with competencies first, motivation second.
 *
 * @param {Array} data
 * @returns {Array} data
 */
function reorderDataset(data) {
  return data.slice(1).map((r) =>
    [r[1]].concat(
      r
        .slice(2)
        .filter((e, i) => i % 2 === 0)
        .concat(r.slice(2).filter((e, i) => i % 2 === 1))
    )
  );
}

/**
 * Formats the sheet headers
 */
function insertGruppeoversiktHeaders(sheet, competencies, motivation, gruppe) {
  const headers = competencies.concat(motivation);

  sheet.getRange(1, 1).setValue(gruppe.visningsnavn).setFontSize(14).setFontWeight('bold');
  sheet
    .getRange(1, 2, 1, headers.length)
    .setValues([headers])
    .setFontSize(10)
    .setFontWeight('bold')
    .setTextRotation(45)
    .setVerticalAlignment('bottom')
    .setHorizontalAlignment('left');
  sheet.setColumnWidths(1, 1, 240);
  sheet.setColumnWidths(2, sheet.getLastColumn() - 1, 30);

  if (sheet.getLastColumn() == sheet.getMaxColumns()) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
  }
  sheet.setColumnWidth(sheet.getMaxColumns(), 240);

  // -------------------------------------------------------------------------
  // Set borders for rotation
  // -------------------------------------------------------------------------
  sheet
    .getRange(1, 2, 1, competencies.length)
    .setBackground(config.colors.competencies.header)
    .setBorder(true, false, false, false, true, false, config.colors.competencies.header, null);
  sheet
    .getRange(1, 2 + competencies.length, 1, motivation.length)
    .setBackground(config.colors.motivation.header)
    .setBorder(true, false, false, false, true, false, config.colors.motivation.header, null);
}

/**
 * Adds the sheet with information about competence and motivaton for group
 * @param {*} ss
 * @param {*} gruppeleder
 */
function addKompetanseVsMotivasjon(ss, gruppeleder) {
  const sheet = ss.getSheetByName('Kompetanse vs Motivasjon') || ss.insertSheet('Kompetanse vs Motivasjon');
  sheet.clear();

  const titlePos = populateKompVsMotCompetencyColumn(sheet);
  const ansattdata = getDataForEmployeesInGroup(gruppeleder.ansatte);
  // console.log("Ansatt nr 0:", gruppeleder.ansatte[0].googleid);
  // console.log("Ansattdata", ansattdata);

  const ansattnavn = gruppeleder.ansatte.map((a) => [a.visningsnavn]).sort();

  formatCompVsMotHeaders(ansattnavn, sheet);

  const comp = ansattdata[gruppeleder.ansatte[0].googleid].competencies.length;
  const mots = ansattdata[gruppeleder.ansatte[0].googleid].motivations.length;

  const go = ss.getSheetByName('Gruppeoversikt');
  const a1comp = go.getRange(2, 2, go.getLastRow() - 1, comp).getA1Notation();
  const a1mots = go.getRange(2, 2 + comp, go.getLastRow(), mots).getA1Notation();

  const a1catHeaders = go.getRange(1, 2, 1, comp).getA1Notation();
  const a1motHeaders = go.getRange(1, 2 + comp, 1, mots).getA1Notation();

  for (let i = 3; i <= sheet.getLastRow(); i++) {
    if (titlePos.includes(i)) {
      const nextValue = titlePos[titlePos.indexOf(i) + 1] || sheet.getLastRow() + 1;
      const rows = nextValue - i - 1;
      for (let j = 2; j < ansattnavn.length + 2; j++) {
        // console.log("Average comp:", i, j, rows);
        const avgCompA1 = sheet.getRange(i + 1, j, rows, 1).getA1Notation();
        const avgMotA1 = sheet.getRange(i + 1, j + ansattnavn.length + 1, rows, 1).getA1Notation();

        sheet
          .getRange(i, j, 1, 1)
          .setValue(`=IFERROR(AVERAGE(${avgCompA1}))`)
          .setNumberFormat('0.0')
          .setBorder(true, false, true, false, null, null);
        sheet
          .getRange(i, j + ansattnavn.length + 1, 1, 1)
          .setValue(`=IFERROR(AVERAGE(${avgMotA1}))`)
          .setNumberFormat('0.0')
          .setBorder(true, false, true, false, null, null);
      }
      continue;
    }
    // row for competencies
    sheet
      .getRange(i, 2)
      .setValue(`=TRANSPOSE(FILTER('Gruppeoversikt'!${a1comp}, 'Gruppeoversikt'!${a1catHeaders} = $A${i}))`);

    // row for motivation
    sheet
      .getRange(i, 2 + ansattnavn.length + 1)
      .setValue(
        `=TRANSPOSE(FILTER('Gruppeoversikt'!${a1mots}, REGEXMATCH('Gruppeoversikt'!${a1motHeaders}, "(?i)"&$A${i})))`
      );
  }

  sheet
    .getRange(3, 2, sheet.getLastRow(), 2 + ansattnavn.length * 2)
    .setFontSize(8)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const compRule = getCompetenciesFormatRule(sheet.getRange(3, 2, sheet.getLastRow(), ansattnavn.length));
  const motsRule = getMotivationsFormatRule(
    sheet.getRange(3, 3 + ansattnavn.length, sheet.getLastRow(), ansattnavn.length)
  );

  let rules = sheet.getConditionalFormatRules();
  rules.push(compRule, motsRule);
  sheet.setConditionalFormatRules(rules);

  sheet.insertColumnAfter(sheet.getLastColumn());
  sheet.setColumnWidth(sheet.getLastColumn() + 1, 300);
  sheet
    .getRange(1, sheet.getLastColumn() + 1, sheet.getLastRow(), 1)
    .setBackground('#ffffff')
    .setFontColor('#ffffff')
    .setValue('x');

  removeUnusedRows(sheet);
  removeUnusedColumns(sheet);
}

/**
 * Adds a sheet with metadata.
 * @param {*} ss
 * @param {*} gruppe
 */
function addInformation(ss, gruppe) {
  const sheet = ss.getSheetByName('Informasjon') || ss.insertSheet('Informasjon');
  sheet.clear();

  sheet.getRange(1, 1, 4, 2).setValues([
    ['Gruppeleder', gruppe.visningsnavn],
    ['Google ID', gruppe.googleid],
    ['Dokument oppdatert', new Date().toJSON()],
    ['Script versjon', `v${config.version}`],
  ]);

  sheet.setColumnWidths(1, 2, 200);
  sheet.getRange(1, 1, sheet.getLastRow(), 1).setFontWeight('bold');

  removeUnusedRows(sheet);
  removeUnusedColumns(sheet);
}

function removeUnusedRows(sheet) {
  if (sheet.getMaxRows() <= sheet.getLastRow()) {
    return;
  }

  sheet.deleteRows(sheet.getLastRow() + 1, sheet.getMaxRows() - sheet.getLastRow());
}

function removeUnusedColumns(sheet) {
  if (sheet.getMaxColumns() <= sheet.getLastColumn()) {
    return;
  }

  sheet.deleteColumns(sheet.getLastColumn() + 1, sheet.getMaxColumns() - sheet.getLastColumn());
}

/**
 * Adds the Gruppeoversikt sheet and populates it with data
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {KIONGroup} gruppe
 */
function addGruppeoversikt(ss, gruppe) {
  const sheetOverview = ss.getSheetByName('Gruppeoversikt') || ss.insertSheet('Gruppeoversikt');

  sheetOverview.clear();

  const sheetSvar = config.sheets.svar;
  const svarData = sheetSvar.getDataRange().getValues();
  const competencies = svarData[0].slice(2).filter((e, i) => i % 2 === 0);
  const motivation = svarData[0].slice(2).filter((e, i) => i % 2 === 1);
  const data = reorderDataset(svarData);

  insertGruppeoversiktHeaders(sheetOverview, competencies, motivation, gruppe);

  // -------------------------------------------------------------------------
  // Process data and populate sheet.
  // -------------------------------------------------------------------------
  const ansatteRows = getDataForAnsatte(gruppe.ansatte, data) || [];
  const missing = getMissingAnsatte(gruppe.ansatte, data) || [];
  const all = ansatteRows
    .concat(missing)
    .map((r) => {
      r[0] = gruppe.ansatte.find((a) => a.googleid === r[0]).visningsnavn;
      return r;
    })
    .sort((a, b) => a[0].localeCompare(b[0]));

  if (all) {
    // console.log("All:", all.length, all[0].length, all[0]);
    sheetOverview
      .getRange(2, 1, all.length, all[0].length)
      .setValues(all)
      .setFontSize(8)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
  }

  // -------------------------------------------------------------------------
  // Fix alignment of first column
  // -------------------------------------------------------------------------
  sheetOverview
    .getRange(2, 1, sheetOverview.getLastRow(), 1)
    .setFontSize(10)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // -------------------------------------------------------------------------
  // Adds conditional format rules to a sheet
  // -------------------------------------------------------------------------
  const catRange = sheetOverview.getRange(2, 2, sheetOverview.getLastRow(), competencies.length);
  const catRule = getCompetenciesFormatRule(catRange);

  const motRange = sheetOverview.getRange(2, 2 + competencies.length, sheetOverview.getLastRow(), motivation.length);
  const motRule = getMotivationsFormatRule(motRange);

  let rules = sheetOverview.getConditionalFormatRules();
  rules.push(catRule, motRule);
  sheetOverview.setConditionalFormatRules(rules);

  removeUnusedRows(sheetOverview);
}

/**
 * Removes a spreadsheet with given name from the document
 *
 * @param {SpreadSheet} ss
 * @param {string} name
 */
function removeSheet(ss, name = 'Sheet1') {
  const dummy = ss.getSheetByName(name);
  if (dummy) {
    ss.deleteSheet(dummy);
  }
}

/**
 * Filters the dataset and returns only the rows that matches the ansatte array
 *
 * @param {Array} ansatte
 * @param {Array} data
 * @returns Array
 */
function getDataForAnsatte(ansatte, data) {
  return data.filter((row) => ansatte.find((a) => a.googleid === row[0]));
}

/**
 * Iterates over the dataset and returns the list of ansatte not found.
 *
 * @param {Array} ansatte
 * @param {Array} data
 * @returns Array
 */
function getMissingAnsatte(ansatte, data) {
  const missing = [];
  ansatte.forEach((a) => {
    if (!data.find((r) => r[0] === a.googleid)) {
      const row = new Array(data[0].length - 1);
      row.unshift(a.googleid);
      missing.push(row);
    }
  });

  // console.log("Missing:", missing);
  return missing;
}

function getMissingEmployees(employees, data) {
  const missing = [];
  employees.forEach((a) => {
    if (!data.find((r) => r[1] === a.googleid)) {
      const row = new Array(data[0].length - 2).fill('');
      row.unshift('', a.googleid);
      missing.push(row);
    }
  });
  return missing;
}

function getDataForEmployeesInGroup(employees) {
  // Get all values and remove timestamp row
  const alldata = config.sheets.svar.getDataRange().getValues().slice(1);

  // console.log(employees);
  // console.log(alldata[0]);
  const missing = getMissingEmployees(employees, alldata);
  // console.log("Missing:", missing);
  const responses = alldata
    .filter((r) => employees.find((a) => a.googleid === r[1]))
    .concat(missing)
    .map((r) => {
      // console.log("Inside map:", r);
      const emp = employees.find((a) => a.googleid === r[1]) || {
        googleid: r[1],
      };
      emp.competencies = r.slice(2).filter((e, i) => i % 2 === 0);
      emp.motivations = r.slice(2).filter((e, i) => i % 2 === 1);

      return emp;
    });

  // console.log(responses[0]);

  data = {};
  responses.forEach((r) => {
    data[r.googleid] = r;
  });
  return data;
}

/**
 * Returns an object with all the competency categories
 */
function getCategories() {
  const sheet = config.sheets.kategorier;

  const data = sheet.getRange(1, 1, sheet.getLastRow(), 7).getValues();

  const titles = data[0];
  const categories = {};

  data.slice(1).forEach((row) => {
    for (i = 0; i < row.length; i++) {
      if (typeof categories[titles[i]] === 'undefined') {
        categories[titles[i]] = [];
      }
      if (typeof row[i] === 'string' && row[i].length > 0) {
        categories[titles[i]].push(row[i]);
      }
    }
  });

  return categories;
}

/**
 * Fetch groups with managers and employees out of the sheet.
 *
 * @returns {Object} List of group managers
 */
function getGroupManagers() {
  const sheet = config.sheets.ansatte;
  const values = sheet.getDataRange().getValues();
  const ledere = getLedereList();
  // console.log("ledere list:", ledere);

  const grupper = ledere.map((item) => {
    item.ansatte = [];
    values.slice(3).forEach((row) => {
      if (row[6] === item.googleid) {
        const ansatt = {
          googleid: row[2],
          lederid: row[6],
          navn: row[0],
          visningsnavn: row[0].split(',').reverse().join(' ').trim(),
          brukernavn: row[3],
        };
        item.ansatte.push(ansatt);
      }
    });
    return item;
  });

  return grupper;
}

// function

/**
 * Fetches the list of group managers from the correct spreadsheet
 * @returns {Array} List of managers
 */
function getLedereList() {
  const sheet = config.sheets.ledere;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn() - 1).getValues();

  const headers = values.shift();

  return values
    .map((row) => {
      let res = {};
      for (i = 0; i < headers.length; i++) {
        res[headers[i].replace(/\s+/g, '').toLowerCase()] = row[i];
      }
      res.visningsnavn = `${res.fornavn} ${res.etternavn}`;
      return res;
    })
    .filter((r) => r.googleid.length > 0);
}

/**
 * Creates the data master spreadsheet - for sharing.
 */
function createOrFetchSpreadsheet(filename) {
  const file = getExistingFile(filename);
  const ss = file !== false ? SpreadsheetApp.open(file) : SpreadsheetApp.create(filename);

  console.log('Got sheet:', ss.getName(), ss.getUrl());
  return ss;
}

/**
 * Searches through the drive looking for existing files with filename. Uses the first found.
 */
function getExistingFile(filename) {
  const files = DriveApp.getFilesByName(filename);

  if (files.hasNext()) {
    return files.next();
  }

  return false;
}

/**
 * Poor mans ID generator. Truncates the HMAC SHA256 signature to 128bits and Mimics an UUID for aesthetics.
 * ID should be consistent through updates though.
 *
 * @param input string
 * @return UUID compatible string.
 */
function getHmacSHA256SigAsUUID(input) {
  const sig = Utilities.computeHmacSha256Signature(input, secrets.hmac_secret)
    .map((chr) => (chr + 256).toString(16).slice(-2))
    .join('');
  return [
    sig.substring(0, 8),
    sig.substring(8, 12),
    sig.substring(12, 16),
    sig.substring(16, 20),
    sig.substring(20, 32),
  ].join('-');
}

/**
 * Formats the headers for the Kompetanse vs Motivasjon sheet.
 *
 * @param {Array} ansattnavn
 * @param {Sheet} sheet
 */
function formatCompVsMotHeaders(ansattnavn, sheet) {
  sheet
    .getRange(1, 2, 1, ansattnavn.length)
    .setValues([ansattnavn])
    .setHorizontalAlignment('left')
    .setVerticalAlignment('bottom')
    .setFontWeight('bold')
    .setTextRotation(45)
    .setBackground(config.colors.competencies.header)
    .setBorder(true, false, false, false, true, false, config.colors.competencies.header, null);

  sheet
    .getRange(1, 2 + ansattnavn.length + 1, 1, ansattnavn.length)
    .setValues([ansattnavn])
    .setHorizontalAlignment('left')
    .setVerticalAlignment('bottom')
    .setFontWeight('bold')
    .setTextRotation(45)
    .setBackground(config.colors.motivation.header)
    .setBorder(true, false, false, false, true, false, config.colors.motivation.header, null);

  sheet
    .getRange(2, 2, 1, ansattnavn.length)
    .merge()
    .setValue('Kompetanse')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setBackground(config.colors.competencies.gradient.max);
  sheet
    .getRange(2, 2 + ansattnavn.length + 1, 1, ansattnavn.length)
    .merge()
    .setValue('Motivasjon')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setBackground(config.colors.motivation.gradient.max);

  // Set width of all columns used to show data.
  sheet.setColumnWidths(2, ansattnavn.length * 2 + 1, 30);
  // resize one spacer column
  sheet.setColumnWidth(2 + ansattnavn.length, 6);
}

/**
 * Calculates the distance between headers in the categories column in
 * the Kompetanse vs Motivasjon sheet.
 *
 * @returns {Array} List of title positions.
 */
function getKompVsMotTitlePosArray(cats) {
  const titlePos = [3];

  Object.keys(cats).forEach((t) => {
    titlePos.push(cats[t].length + titlePos.slice(-1)[0] + 1);
  });
  titlePos.pop();

  return titlePos;
}
/**
 * Populates the Competency column of the Kompetanse vs Motivaition Sheet.
 *
 * @param {Sheet} sheet
 * @returns {Array} the list of title positions in the competency column
 */
function populateKompVsMotCompetencyColumn(sheet) {
  const categories = getCategories();
  const titlePos = getKompVsMotTitlePosArray(categories);

  const data = Object.keys(categories)
    .map((k) => [k].concat(categories[k]))
    .flat()
    .map((i) => [i]);

  sheet.getRange(titlePos[0], 1, data.length, 1).setValues(data);
  titlePos.forEach((pos) => {
    sheet.getRange(pos, 1).setFontWeight('bold').setBackground('#efefef');
  });

  sheet.setColumnWidth(1, 300);

  return titlePos;
}

/**
 * Returns the standard rule for highlighting competencies
 *
 * @param {Range} range to highlight competencies for
 * @return {ConditionalFormatRuleBuilder}
 */
function getCompetenciesFormatRule(range) {
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(config.colors.competencies.gradient.min)
    .setGradientMaxpoint(config.colors.competencies.gradient.max)
    .setRanges([range])
    .build();
}

/**
 * Returns the standard rule for highlighting motivations
 *
 * @param {Range} range to highlight motivations for
 * @return {ConditionalFormatRuleBuilder}
 */
function getMotivationsFormatRule(range) {
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(config.colors.motivation.gradient.min)
    .setGradientMaxpoint(config.colors.motivation.gradient.max)
    .setRanges([range])
    .build();
}
