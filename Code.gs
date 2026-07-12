/**
 * COFFEE DATA — Apps Script backend
 * ---------------------------------
 * Lives inside the "Coffee Data" Google Sheet (Extensions → Apps Script).
 *
 * Functions:
 *   setup()  — run ONCE by hand. Builds the Roast Log, Espresso Log, Beans,
 *              and Settings tabs with headers + formatting. Safe to re-run:
 *              it never deletes data, only adds missing tabs/headers.
 *   doPost() — receives a JSON entry from the capture app and appends a row.
 *   doGet()  — returns bean list, roast lots, and custom-field settings
 *              so the app can populate its dropdowns.
 */

// ---------- column definitions (order = sheet column order) ----------

var ROAST_HEADERS = [
  'Active', 'Roast Date', 'Bean', 'Variant', 'Lot Number',
  'Start Weight (g)', 'Preheat P1 (s)', 'Behmor Profile',
  'First Crack (M:SS)', 'Total Roast (M:SS)', 'Development (M:SS)',
  'End Weight (g)', 'Weight Loss (%)', 'Roast Level (Agtron)',
  'Notes', 'Custom 1', 'Custom 2', 'Custom 3', 'Custom 4', 'Logged'
];

var ESPRESSO_HEADERS = [
  'Brew Date', 'Bean / Lot', 'Bean Source', 'Pourover (Y/N)',
  'Grind Level', 'Bean Weight (g)', 'Pre-infusion (s)', 'Brew Time (s)',
  'Yield (g)', 'Ratio', 'Days Rested',
  'Likeability (L)', 'Crema (T)', 'Acidity (T)', 'Bitterness (T)',
  'Body (T)', 'Roastiness (T)', 'Medicinal (T)',
  'Floral (T)', 'Honey (T)', 'Sugars (T)', 'Caramel (T)', 'Fruits (T)',
  'Citrus (T)', 'Berry (T)', 'Cocoa (T)', 'Nuts (T)', 'Rustic (T)', 'Spice (T)',
  'Notes', 'Basket', 'Spring (psi)', 'Puck Screen (Y/N)',
  'Custom 1', 'Custom 2', 'Custom 3', 'Custom 4', 'Logged'
];

// keys the app sends, in the same order as the headers above
var ROAST_KEYS = [
  'active', 'roastDate', 'bean', 'variant', 'lot',
  'startWeight', 'preheat', 'profile',
  'firstCrack', 'totalRoast', 'development',
  'endWeight', 'weightLoss', 'roastLevel',
  'notes', 'custom1', 'custom2', 'custom3', 'custom4'
];

var ESPRESSO_KEYS = [
  'brewDate', 'bean', 'source', 'pourover',
  'grind', 'dose', 'preinfusion', 'brewTime',
  'yield', 'ratio', 'daysRested',
  'likeability', 'crema', 'acidity', 'bitterness',
  'body', 'roastiness', 'medicinal',
  'floral', 'honey', 'sugars', 'caramel', 'fruits',
  'citrus', 'berry', 'cocoa', 'nuts', 'rustic', 'spice',
  'notes', 'basket', 'spring', 'puckScreen',
  'custom1', 'custom2', 'custom3', 'custom4'
];

var SEED_BEANS = [
  ['India Pearl Mountain Estate Peaberry', 'Washed', 'Sweet Maria\'s', 'Active', 'Cocoa, malted grain. City+ to Full City+. Espresso: yes'],
  ['Burundi Colline Rugembe', 'Washed', 'Sweet Maria\'s', 'Active', 'Dark choc truffle, caramel. City+ to Full City+. Espresso: yes'],
  ['India Chikmagalur Gemberly Robusta', 'Washed', 'Sweet Maria\'s', 'Active', 'Blend component ~25%. Full City to French'],
  ['Peru Comite La Palma', 'Washed', 'Sweet Maria\'s', 'Active', ''],
  ['Java Sunda Wet Hulled Rancabali', 'Wet Hulled', 'Sweet Maria\'s', 'Active', 'P1, preheat, manual at P5. Full City first'],
  ['Rwanda Nyamasheke Coproca', 'Washed', 'Sweet Maria\'s', 'Active', ''],
  ['Costa Rica Anaerobic Higuito', 'Anaerobic', 'Sweet Maria\'s', 'Active', 'Rest 12-14 days before espresso'],
  ['Guatemala El Socorro Java', 'Washed', 'Sweet Maria\'s', 'Active', '']
];

// ---------- one-time setup ----------

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  makeLogTab(ss, 'Roast Log', ROAST_HEADERS, '#586B4E');
  makeLogTab(ss, 'Espresso Log', ESPRESSO_HEADERS, '#B37F2B');
  makeBeansTab(ss);
  makeSettingsTab(ss);

  // remove the default empty Sheet1 if it exists and is empty
  var s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

function makeLogTab(ss, name, headers, color) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(color).setFontColor('#FFFFFF')
      .setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, headers.length, 110);
    sh.setColumnWidth(headers.indexOf('Notes') + 1, 260);
    sh.setTabColor(color);
  }
}

function makeBeansTab(ss) {
  var sh = ss.getSheetByName('Beans') || ss.insertSheet('Beans');
  if (sh.getLastRow() === 0) {
    var headers = ['Bean', 'Process', 'Source', 'Status', 'Notes'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#4A3B2E').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, SEED_BEANS.length, 5).setValues(SEED_BEANS);
    // Status dropdown
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Finished'], true).build();
    sh.getRange(2, 4, 500, 1).setDataValidation(rule);
    sh.setColumnWidth(1, 280);
    sh.setColumnWidth(5, 320);
    sh.setTabColor('#4A3B2E');
  }
}

function makeSettingsTab(ss) {
  var sh = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['Field', 'Custom Name', 'Active (x)']])
      .setFontWeight('bold').setBackground('#666').setFontColor('#FFF');
    var rows = [
      ['Roast Custom 1', '', ''], ['Roast Custom 2', '', ''],
      ['Roast Custom 3', '', ''], ['Roast Custom 4', '', ''],
      ['Espresso Custom 1', '', ''], ['Espresso Custom 2', '', ''],
      ['Espresso Custom 3', '', ''], ['Espresso Custom 4', '', '']
    ];
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 200);
    sh.getRange('E1').setValue('Name a custom slot and put x under Active — the app shows it on the next refresh.');
  }
}

// ---------- web app endpoints ----------

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh, keys;

    if (data.type === 'roast') {
      sh = ss.getSheetByName('Roast Log');
      keys = ROAST_KEYS;
    } else if (data.type === 'espresso') {
      sh = ss.getSheetByName('Espresso Log');
      keys = ESPRESSO_KEYS;
    } else {
      return jsonOut({ ok: false, error: 'Unknown entry type' });
    }

    var row = keys.map(function (k) {
      var v = data[k];
      return (v === undefined || v === null) ? '' : v;
    });
    row.push(new Date()); // Logged timestamp
    sh.appendRow(row);

    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Active beans
    var beans = [];
    var bSheet = ss.getSheetByName('Beans');
    if (bSheet && bSheet.getLastRow() > 1) {
      var bVals = bSheet.getRange(2, 1, bSheet.getLastRow() - 1, 4).getValues();
      bVals.forEach(function (r) {
        if (r[0] && String(r[3]).toLowerCase() !== 'finished') {
          beans.push({ name: String(r[0]), process: String(r[1]) });
        }
      });
    }

    // Roast lots (most recent 40, Active=Yes only) with roast dates,
    // for the espresso dropdown + days-rested calc
    var lots = [];
    var rSheet = ss.getSheetByName('Roast Log');
    if (rSheet && rSheet.getLastRow() > 1) {
      var start = Math.max(2, rSheet.getLastRow() - 79); // scan more rows since some get filtered out
      var rVals = rSheet.getRange(start, 1, rSheet.getLastRow() - start + 1, 5).getValues();
      rVals.forEach(function (r) {
        var active = String(r[0]).trim().toLowerCase();
        var lot = r[4];
        if (lot && active !== 'no') { // blank/Yes/anything but explicit "No" counts as active
          lots.push({
            lot: String(lot),
            bean: String(r[2]),
            roastDate: r[1] instanceof Date
              ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'M/d/yyyy')
              : String(r[1])
          });
        }
      });
      lots.reverse(); // newest first
      lots = lots.slice(0, 40);
    }

    // Custom field settings
    var custom = { roast: [], espresso: [] };
    var sSheet = ss.getSheetByName('Settings');
    if (sSheet && sSheet.getLastRow() > 1) {
      var sVals = sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 3).getValues();
      sVals.forEach(function (r) {
        var field = String(r[0]);
        var entry = {
          name: String(r[1] || ''),
          active: String(r[2]).trim().toLowerCase() === 'x' && !!r[1]
        };
        if (field.indexOf('Roast') === 0) custom.roast.push(entry);
        else if (field.indexOf('Espresso') === 0) custom.espresso.push(entry);
      });
    }

    return jsonOut({ ok: true, beans: beans, lots: lots, custom: custom });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
