/** create basic layout for the Tab Statistics
 *
 */
function createStatisticsLayout() {
  if (statisticsSheet == null) {
    ss.insertSheet("Statistics", 0);
    statisticsSheet = ss.getSheetByName("Statistics");
  }
  var statisticsRange = statisticsSheet.getRange(1, 1, 3, 17),
    statisticsValue = statisticsRange.getValues();

  statisticsValue[1][0] = "Total Count of valid Logs:";
  statisticsValue[1][1] = "Participation";
  statisticsValue[1][3] = "First Death";
  statisticsValue[1][5] = "Downs";
  statisticsValue[1][7] = "Res";
  statisticsValue[1][9] = "Deads";
  statisticsValue[1][11] = "ResTime";
  statisticsValue[1][12] = "damageTaken";
  statisticsValue[1][13] = "DPS";
  statisticsValue[1][14] = "Breakbar";
  statisticsValue[1][15] = "CondiCleanses";
  statisticsValue[1][16] = "BoonStrips";

  statisticsValue[2][0] = "=COUNTA(Logs!B2:B)";
  for (i = 0; i < 5; i++) {
    statisticsValue[2][i * 2 + 1] = "total";
    statisticsValue[2][i * 2 + 2] = "percent";
  }
  for (i = 0; i < 6; i++) {
    statisticsValue[2][i + 11] = "AVG";
  }

  for (i = 0; i < 5; i++) {
    statisticsSheet.getRange(2, i * 2 + 2, 1, 2).mergeAcross();
  }

  statisticsRange
    .setValues(statisticsValue)
    .setFontFamily("Arial")
    .setFontSize(11)
    .setFontWeight("bold");
  statisticsSheet.getRange(2, 1, 2, 17).setHorizontalAlignment("center");
  statisticsSheet
    .getRange(2, 1, 2, 17)
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

  statisticsSheet.getRange(2, 1, 1, 17).setBackground(gray);
  statisticsSheet.autoResizeColumns(1, 1).setColumnWidths(2, 10, 60);

  if (statisticsSheet.getLastRow() < 40 && statisticsSheet.getMaxRows() != 40) {
    statisticsSheet.deleteRows(40, statisticsSheet.getMaxRows() - 40);
  }
  if (
    statisticsSheet.getLastColumn() < 18 &&
    statisticsSheet.getMaxColumns() > 18
  ) {
    statisticsSheet.deleteColumns(18, statisticsSheet.getMaxColumns() - 18);
  }

  var statisticsProtection = statisticsSheet.protect(),
    me = Session.getEffectiveUser();

  statisticsProtection
    .removeEditors(statisticsProtection.getEditors())
    .addEditor(me);

  var triggers = ScriptApp.getProjectTriggers();
  if (
    !triggers.some(
      (trigger) => trigger.getHandlerFunction() == "updateStatisticsTrigger"
    )
  ) {
    ScriptApp.newTrigger("updateStatisticsTrigger")
      .forSpreadsheet(ss)
      .onChange()
      .create();
  }
}

/** create basic layout for the Tab Statistics
 *
 */
function cleanUpStatisticsLayout() {
  var statisticsRange = statisticsSheet.getRange(
    3,
    1,
    statisticsSheet.getLastRow() - 2,
    18
  ),
    statisticsValue = statisticsRange.getValues();

  for (var i = 0; i < statisticsValue.length; i++) {
    statisticsSheet
      .getRange(i + 4, 1, 1, 17)
      .setValues([
        ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
      ])
      .setBorder(false, false, false, false, false, false);
  }
  statisticsSheet
    .getRange(4, 1, statisticsSheet.getLastRow() - 3, 17)
    .setBorder(
      null,
      null,
      true,
      null,
      null,
      null,
      "black",
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  updateStatisticsLayout(statisticsValue.length - 1);
}

/** 
 */
function updateStatisticsLayout(amountOfPlayers) {
  if (amountOfPlayers > statisticsSheet.getMaxRows() - 10) {
    statisticsSheet.insertRowsAfter(
      statisticsSheet.getMaxRows(),
      amountOfPlayers - (statisticsSheet.getMaxRows() - 10)
    );
  }
  var rules = new Array();
  // Layout settings for the list of players including the Participation and first Deaths
  for (i = 0; i < 6; i++) {
    statisticsSheet
      .getRange(4, i * 2 + 1, amountOfPlayers, 1)
      .setBorder(
        null,
        null,
        null,
        true,
        null,
        null,
        "black",
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
    statisticsSheet
      .getRange(4, i * 2 + 2, amountOfPlayers, 1)
      .setNumberFormat("#,##0");
  }

  for (i = 0; i < 5; i++) {
    statisticsSheet
      .getRange(4, i * 2 + 3, amountOfPlayers, 1)
      .setNumberFormat("#0.00%");
  }

  statisticsSheet
    .getRange(4, 12, amountOfPlayers, 6)
    .setNumberFormat("#,##0.0")
    .setBorder(
      null,
      null,
      null,
      null,
      true,
      null,
      "black",
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );

  var participationRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([statisticsSheet.getRange(4, 3, amountOfPlayers, 1)])
    .build();
  rules.push(participationRule);

  var firstDeathRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(red)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(green)
    .setRanges([statisticsSheet.getRange(4, 5, amountOfPlayers, 1)])
    .build();
  rules.push(firstDeathRule);

  var downRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(red)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(green)
    .setRanges([statisticsSheet.getRange(4, 7, amountOfPlayers, 1)])
    .build();
  rules.push(downRule);

  var resRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([statisticsSheet.getRange(4, 9, amountOfPlayers, 1)])
    .build();
  rules.push(resRule);

  var deadRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(red)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(green)
    .setRanges([statisticsSheet.getRange(4, 11, amountOfPlayers, 1)])
    .build();
  rules.push(deadRule);

  var resTimeRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([statisticsSheet.getRange(4, 12, amountOfPlayers, 1)])
    .build();
  rules.push(resTimeRule);

  var dmgTakenRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(red)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(green)
    .setRanges([statisticsSheet.getRange(4, 13, amountOfPlayers, 1)])
    .build();
  rules.push(dmgTakenRule);

  var dpsRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([statisticsSheet.getRange(4, 14, amountOfPlayers, 1)])
    .build();
  rules.push(dpsRule);

  var breakbarDamageRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([statisticsSheet.getRange(4, 15, amountOfPlayers, 1)])
    .build();
  rules.push(breakbarDamageRule);

  var condiCleanseRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([statisticsSheet.getRange(4, 16, amountOfPlayers, 1)])
    .build();
  rules.push(condiCleanseRule);

  var boonStripsRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([statisticsSheet.getRange(4, 17, amountOfPlayers, 1)])
    .build();
  rules.push(boonStripsRule);

  statisticsSheet.setConditionalFormatRules(rules);
}
