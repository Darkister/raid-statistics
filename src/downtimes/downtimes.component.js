/** Update the Statistics when ever there is a change on the logs sheet
 * @param {*} e
 */
function updateDowntimeTrigger(e) {
    Logger.log(e.changeType);
    if (
        e &&
        e.source &&
        (e.changeType === "REMOVE_ROW" ||
            e.changeType === "EDIT" ||
            e.changeType === "OTHER")
    ) {
        downtimeSheet
            .getRange(4, 1, downtimeSheet.getMaxRows() - 4, 14)
            .clear();
        var amountOfDays = fillAllDays();
        console.log(amountOfDays);
        updateDowntimeLayout(amountOfDays);
    }
}

/**
 * Get all different Days in the logs tab
 */
function fillAllDays() {
    var days = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues(),
        flatDays = days.flat(),
        uniqueDates = new Set(flatDays.map(date => date.getTime())), // Vergleicht die Zeitstempel
        uniqueDateObjects = Array.from(uniqueDates).map(timestamp => new Date(timestamp));

    Logger.log(uniqueDateObjects.length);

    var fillDays = downtimeSheet.getRange(4, 1, uniqueDateObjects.length, 14),
        daysValues = new Array(uniqueDateObjects.length).fill().map((_) => []);

    Logger.log(daysValues);
    for (var a = 0; a < daysValues.length; a++) {
        daysValues[a].push(
            // Date
            uniqueDateObjects[a],
            // Encounters
            "=COUNTIF(Logs!A2:A;A" + (a + 4) + ")",
            // Kills
            "=COUNTIFS(Logs!A2:A;A" + (a + 4) + ";Logs!D2:D;TRUE)",
            // Successrate
            "=C" + (a + 4) + "/B" + (a + 4),
            // Raid Start
            "=MINIFS(Logs!H2:H;Logs!A2:A;A" + (a + 4) + ") + (6/24)",
            // Raid End
            "=MAXIFS(Logs!I2:I;Logs!A2:A;A" + (a + 4) + ") + (6/24)",
            // Duration
            "=F" + (a + 4) + "-E" + (a + 4),
            // Overall Encounter Time
            "=SUMIF(Logs!A2:A;A" + (a + 4) + ";Logs!G2:G) / 1000 / 60 / 60 / 24",
            // Overall Encounter %
            "=H" + (a + 4) + "/G" + (a + 4),
            // Overall Downtime
            "=G" + (a + 4) + "-H" + (a + 4),
            // fill
            "",
            // negativ fails Encounter Time
            "=SUMIFS(Logs!G2:G;Logs!A2:A;A" + (a + 4) + ";Logs!D2:D;TRUE) / 1000 / 60 / 60 / 24",
            // negativ fails Encounter %
            "=L" + (a + 4) + "/G" + (a + 4),
            // negativ fails Downtime
            "=G" + (a + 4) + "-L" + (a + 4)
        );
    }

    fillDays
        .setValues(daysValues)
        .setBorder(false, false, false, false, false, false)
        .setFontSize(11)
        .setFontFamily("Arial")
        .setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
    downtimeSheet.getRange(1, 11, 3 + uniqueDateObjects.length, 1).setBackground(black);

    return uniqueDateObjects.length;
}
