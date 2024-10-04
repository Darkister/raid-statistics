/** create basic layout for the tab Up-/Downtime
 *
 */
function createDowntimeLayout() {
    if (downtimeSheet == null) {
        ss.insertSheet("Up-/Downtime", 1);
        downtimeSheet = ss.getSheetByName("Up-/Downtime");
    }

    var downtimeRange = downtimeSheet.getRange(1, 1, 3, 14),
        downtimeValue = downtimeRange.getValues();

    downtimeValue[0][7] = "Overall";
    downtimeValue[0][11] = "negativ fails";
    downtimeValue[1][0] = "Date";
    downtimeValue[1][1] = "Encounters";
    downtimeValue[1][2] = "Kills";
    downtimeValue[1][3] = "Successrate";
    downtimeValue[1][4] = "Raid Start";
    downtimeValue[1][5] = "Raid End";
    downtimeValue[1][6] = "Duration";
    downtimeValue[1][7] = "Encounter Time";
    downtimeValue[1][8] = "Encounter %";
    downtimeValue[1][9] = "Downtime";
    downtimeValue[1][11] = "Encounter Time";
    downtimeValue[1][12] = "Encounter %";
    downtimeValue[1][13] = "Downtime";
    downtimeValue[2][0] = "AVG";
    downtimeValue[2][1] = "=AVERAGE(B4:B)";
    downtimeValue[2][2] = "=AVERAGE(C4:C)";
    downtimeValue[2][3] = "=AVERAGE(D4:D)";
    downtimeValue[2][4] = "=AVERAGE(E4:E)";
    downtimeValue[2][5] = "=AVERAGE(F4:F)";
    downtimeValue[2][6] = "=AVERAGE(G4:G)";
    downtimeValue[2][7] = "=AVERAGE(H4:H)";
    downtimeValue[2][8] = "=H3/G3";
    downtimeValue[2][9] = "=AVERAGE(J4:J)";
    downtimeValue[2][11] = "=AVERAGE(L4:L)";
    downtimeValue[2][12] = "=L3/G3";
    downtimeValue[2][13] = "=AVERAGE(N4:N)";

    downtimeRange
        .setValues(downtimeValue)
        .setFontFamily("Arial")
        .setFontSize(11)
        .setFontWeight("bold");
    downtimeSheet.getRange(1, 8, 1, 3).mergeAcross();
    downtimeSheet.getRange(1, 12, 1, 3).mergeAcross();
    downtimeSheet.getRange(2, 1, 1, 14).setBackground(gray);
    downtimeSheet
        .getRange(2, 1, 2, 14)
        .setBorder(true,
            true,
            true,
            true,
            true,
            true,
            "black",
            SpreadsheetApp.BorderStyle.SOLID_THICK
        );
    downtimeSheet.getRange(1, 1, 3, 14).setHorizontalAlignment("center");
    downtimeSheet.setColumnWidth(1, 82);
    downtimeSheet.setColumnWidths(2, 3, 95);
    downtimeSheet.setColumnWidths(5, 3, 78);
    downtimeSheet.setColumnWidth(8, 120);
    downtimeSheet.setColumnWidth(9, 99);
    downtimeSheet.setColumnWidth(10, 79);
    downtimeSheet.setColumnWidth(11, 25);
    downtimeSheet.setColumnWidth(12, 120);
    downtimeSheet.setColumnWidth(13, 99);
    downtimeSheet.setColumnWidth(14, 79);
    downtimeSheet.getRange(1, 11, 3, 1).setBackground(black);

    if (downtimeSheet.getLastRow() < 20 && downtimeSheet.getMaxRows() != 20) {
        downtimeSheet.deleteRows(20, downtimeSheet.getMaxRows() - 20)
    }
    if (downtimeSheet.getLastColumn() < 15 && downtimeSheet.getMaxColumns() != 15) {
        downtimeSheet.deleteColumns(15, downtimeSheet.getMaxColumns() - 15)
    }

    if (
        !triggers.some((trigger) => trigger.getHandlerFunction() == "updateDowntimeTrigger")
    ) {
        ScriptApp.newTrigger("updateDowntimeTrigger").forSpreadsheet(ss).onEdit().create();
    }
}

/**
 * 
 */
function updateDowntimeLayout(amountOfDays) {
    if (amountOfDays > downtimeSheet.getMaxRows() - 10) {
        downtimeSheet.insertRowsAfter(
            downtimeSheet.getMaxRows(),
            amountOfDays - (downtimeSheet.getMaxRows() - 10)
        );
    }
    var rules = new Array();

    downtimeSheet
        .getRange(3, 4, 1 + amountOfDays, 1)
        .setNumberFormat("#0.00%");

    downtimeSheet
        .getRange(3, 5, 1 + amountOfDays, 2)
        .setNumberFormat("HH:mm:ss");

    downtimeSheet
        .getRange(3, 7, 1 + amountOfDays, 2)
        .setNumberFormat("[h]:mm:ss");

    downtimeSheet
        .getRange(3, 9, 1 + amountOfDays, 1)
        .setNumberFormat("#0.00%");

    downtimeSheet
        .getRange(3, 12, 1 + amountOfDays, 1)
        .setNumberFormat("[h]:mm:ss");

    downtimeSheet
        .getRange(3, 13, 1 + amountOfDays, 1)
        .setNumberFormat("#0.00%");

    var successrateRule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint(green)
        .setGradientMidpointWithValue(
            yellow,
            SpreadsheetApp.InterpolationType.PERCENTILE,
            "50"
        )
        .setGradientMinpoint(red)
        .setRanges([downtimeSheet.getRange(4, 4, amountOfDays, 1)])
        .build();
    rules.push(successrateRule);

    var overallEncounterPercent = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint(green)
        .setGradientMidpointWithValue(
            yellow,
            SpreadsheetApp.InterpolationType.PERCENTILE,
            "50"
        )
        .setGradientMinpoint(red)
        .setRanges([downtimeSheet.getRange(4, 9, amountOfDays, 1)])
        .build();
    rules.push(overallEncounterPercent);

    var negativfailsEncounterPercent = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint(green)
        .setGradientMidpointWithValue(
            yellow,
            SpreadsheetApp.InterpolationType.PERCENTILE,
            "50"
        )
        .setGradientMinpoint(red)
        .setRanges([downtimeSheet.getRange(4, 13, amountOfDays, 1)])
        .build();
    rules.push(negativfailsEncounterPercent);

    downtimeSheet.setConditionalFormatRules(rules);
}