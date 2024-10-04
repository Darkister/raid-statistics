/** create basic layout for the tab Logs
 *
 */
function createLogsLayout() {
  if (logSheet == null) {
    ss.insertSheet("Logs", 2);
    logSheet = ss.getSheetByName("Logs");
  }
  while (logSheet.getMaxColumns() < 86) {
    logSheet.insertColumnAfter(24);
  }
  var logRange = logSheet.getRange(1, 1, 1, 86),
    logValue = logRange.getValues();

  logValue[0][0] = "Date";
  logValue[0][1] = "Log";
  logValue[0][2] = "Boss/Encounter";
  logValue[0][3] = "Success";
  logValue[0][4] = "Rest HP";
  logValue[0][5] = "Duration";
  logValue[0][6] = "DurationMS";
  logValue[0][7] = "timeStart";
  logValue[0][8] = "timeEnd";
  logValue[0][9] = "CM?";
  logValue[0][10] = "First Death";
  logValue[0][11] = "all Player down on First Death";
  logValue[0][12] = "Players Accountname";
  logValue[0][22] = "FullFight DPS";
  logValue[0][32] = "Breakbar";
  logValue[0][42] = "Received Damage";
  logValue[0][52] = "ResDuration";
  logValue[0][62] = "Condis Cleansed";
  logValue[0][72] = "Boon Strips";
  logValue[0][82] = "Downed";
  logValue[0][83] = "Dead";
  logValue[0][84] = "Got up";
  logValue[0][85] = "Res";

  for (var i = 1; i <= 7; i++) {
    logSheet.getRange(1, i * 10 + 3, 1, 10).mergeAcross();
  }
  logSheet
    .getRange(2, 1, logSheet.getMaxRows() - 1, 1)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold");
  logSheet
    .getRange(2, 3, logSheet.getMaxRows() - 1, 16)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  logSheet
    .getRange(2, 5, logSheet.getMaxRows() - 1, 1)
    .setNumberFormat("#0.00%");
  logRange
    .setValues(logValue)
    .setFontFamily("Arial")
    .setFontSize(11)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "black",
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );

  var filter = logSheet.getFilter();
  if (!filter) {
    logRange.createFilter();
  }

  logSheet
    .setColumnWidths(1, 1, 80)  // Date
    .setColumnWidths(2, 1, 300) // Log
    .setColumnWidths(3, 1, 150) // Boss/Encounter
    .setColumnWidths(4, 3, 85)  // Success | Rest HP | Duration
    .setColumnWidths(7, 1, 150) // DurationMS
    .setColumnWidths(8, 2, 125) // timeStart | timeEnd
    .setColumnWidths(10, 1, 85) // CM?
    .setColumnWidths(11, 1, 100)// First Death
    .setColumnWidths(12, 1, 85) // all Player down on First Death
    .setColumnWidths(13, 10, 25)// Players Accountname
    .setColumnWidths(23, 10, 25)// FullFight DPS
    .setColumnWidths(33, 10, 25)// Breakbar
    .setColumnWidths(43, 10, 25)// Received Damage
    .setColumnWidths(53, 10, 25)// ResDuration
    .setColumnWidths(63, 10, 25)// Condis Cleansed
    .setColumnWidths(73, 10, 25)// Boon Strips
    .setColumnWidths(83, 4, 100)// Downed | Dead | Got up | res
    .setFrozenRows(1);
}

function rebuildFilter() {
  var startRow = 2,
    startColumn = 1,
    lastRow = logSheet.getLastRow(),
    lastColumn = logSheet.getLastColumn(),
    range = logSheet.getRange(
      startRow,
      startColumn,
      lastRow - startRow + 1,
      lastColumn
    ),
    filter = logSheet.getFilter(),
    criteria = [];

  // Store the filter criteria before removing the filter
  if (filter) {
    //var numColumns = range.getNumColumns();
    for (var col = 1; col <= filter.getRange().getLastColumn(); col++) {
      criteria.push(filter.getColumnFilterCriteria(col));
    }
  }

  // Remove the filter
  if (filter) {
    filter.remove();
  }

  // Sort the data
  range.sort([{ column: 1 }]);

  // Reapply the filter
  if (criteria.length > 0) {
    var newFilterRange = logSheet.getRange(
      startRow - 1,
      startColumn,
      lastRow,
      lastColumn
    );
    newFilterRange.createFilter();
    var newFilter = newFilterRange.getFilter();
    var numColumns = newFilterRange.getNumColumns();
    for (var col = 1; col <= numColumns; col++) {
      if (criteria[col - 1]) {
        newFilter.setColumnFilterCriteria(col, criteria[col - 1]);
      }
    }
  }
}
