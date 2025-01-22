/** Update the Statistics when ever there is a change on the logs sheet
 * @param {*} e
 */
function updateStatisticsTrigger(e) {
  if (
    e &&
    e.source &&
    (e.changeType === "REMOVE_ROW" ||
      e.changeType === "EDIT" ||
      e.changeType === "OTHER")
  ) {
    statisticsSheet
      .getRange(4, 1, statisticsSheet.getMaxRows() - 4, 17)
      .clear();
    var amountOfPlayers = fillAllPlayersAccName();
    console.log(amountOfPlayers);
    updateStatisticsLayout(amountOfPlayers);
  }
}

/** Get the Accountnames of all Players and fill it into the Statisticssheet
 */
function fillAllPlayersAccName() {
  var static = staticSheet.getRange(2, 2, staticSheet.getLastRow() - 1, 1).getValues(),
    players = logSheet
      .getRange(2, 13, logSheet.getLastRow() - 1, 10)
      .getValues(),
    allPlayers = new Set();

  // add all unique members of the static
  static.forEach((p) => { if (p != "") { allPlayers.add(p[0]) } });

  // add all other players | add only unique players
  players.forEach((r) => {
    r.forEach((p) => { if (p != "") { allPlayers.add(p) } });
  });

  var fillPlayers = statisticsSheet.getRange(4, 1, allPlayers.size, 17),
    arr = Array.from(allPlayers),
    playerValues = new Array(arr.length).fill().map((_) => []);
  for (var a = 0; a < playerValues.length; a++) {
    playerValues[a].push(
      // Players
      arr[a],
      // Participation total
      "=IFERROR(COUNTIF(Logs!M2:V;A" + (a + 4) + "))",
      // Participation percent
      "=IFERROR(B" + (a + 4) + "/A3)",
      // First Death total
      "=IFERROR(COUNTIFS(Logs!K2:K;A" + (a + 4) + ";Logs!L2:L;FALSE))",
      // First Death percent
      "=IFERROR(D" + (a + 4) + "/B" + (a + 4) + ")",
      // Downs total
      '=IFERROR(COUNTIFS(Logs!CE2:CE;"*" & A' + (a + 4) + ' & "*";Logs!L2:L;FALSE))',
      // Downs percent
      "=IFERROR(F" + (a + 4) + "/B" + (a + 4) + ")",
      // Res total
      '=IFERROR(COUNTIFS(Logs!CH2:CH;"*" & A' + (a + 4) + ' & "*"))',
      // Res percent
      "=IFERROR(H" + (a + 4) + "/B" + (a + 4) + ")",
      // Deads total
      '=IFERROR(COUNTIFS(Logs!CF2:CF;"*" & A' + (a + 4) + ' & "*"))',
      // Deads percent
      "=IFERROR(J" + (a + 4) + "/B" + (a + 4) + ")",
      // Res Duration Average
      "=IFERROR(SUMIFS(Logs!BA2:BJ;Logs!M2:V;A" + (a + 4) + ") / H" + (a + 4) + ")",
      // Damage Taken Average
      "=IFERROR(SUMIFS(Logs!AQ2:AZ;Logs!M2:V;A" + (a + 4) + ") / B" + (a + 4) + ")",
      // DPS Average
      "=IFERROR(SUMIFS(Logs!W2:AF;Logs!M2:V;A" + (a + 4) + ") / B" + (a + 4) + ")",
      // Breakbar Average
      "=IFERROR(SUMIFS(Logs!AG2:AP;Logs!M2:V;A" + (a + 4) + ") / B" + (a + 4) + ")",
      // Condi Cleanses Average
      "=IFERROR(SUMIFS(Logs!BK2:BT;Logs!M2:V;A" + (a + 4) + ") / B" + (a + 4) + ")",
      // Boon Strips Average
      "=IFERROR(SUMIFS(Logs!BU2:CD;Logs!M2:V;A" + (a + 4) + ") / B" + (a + 4) + ")"
    );
  }

  fillPlayers
    .setValues(playerValues)
    .setBorder(false, false, false, false, false, false)
    .setHorizontalAlignment("center")
    .setFontSize(11)
    .setFontFamily("Arial")
    .setBorder(
      true,
      true,
      true,
      true,
      null,
      null,
      "black",
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  statisticsSheet
    .getRange(4, 1, staticSheet.getRange(2, 2, 5, 1).getValues().flat().filter(value => value !== "").length, 17)
    .setBorder(
      null,
      null,
      true,
      null,
      null,
      null,
      "black",
      SpreadsheetApp.BorderStyle.SOLID
    );
  statisticsSheet
    .getRange(4, 1, staticSheet.getRange(2, 2, 10, 1).getValues().flat().filter(value => value !== "").length, 17)
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

  return allPlayers.size;
}
