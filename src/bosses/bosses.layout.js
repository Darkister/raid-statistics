/**
 * create a basic layout for a given Encounter
 * @param {string} encounter Name of Encounter to create a spreadsheet
 * @param {[string]} bossMechanics Array of Mechanics for the given Boss
 */
function createBossSpecificLayout(encounter, bossMechanics) {
  if (ss.getSheetByName(encounter) == null) {
    ss.insertSheet(encounter);
    bossSpecificSheet = ss.getSheetByName(encounter);
  } else {
    bossSpecificSheet = ss.getSheetByName(encounter);
  }
  var bossRange = bossSpecificSheet.getRange(1, 1, 37, 22),
    bossValue = bossRange.getValues(),
    rules = new Array();

  bossValue[0][0] = encounter;

  for (i = 0; i < 10; i++) {
    bossValue[i + 3][0] = "='Setup und Co'!B" + (i + 2);
  }

  for (i = 0; i < 10; i++) {
    bossValue[i + 17][0] = "='Setup und Co'!B" + (i + 2);
  }
  bossValue[1][1] = "Participation";
  bossValue[1][3] = "Kills";
  bossValue[1][5] = "First Death";
  bossValue[1][7] = "Downs";
  bossValue[1][9] = "Res";
  bossValue[1][11] = "Deads";
  bossValue[1][13] = "ResTime";
  bossValue[1][14] = "damageTaken";
  bossValue[1][16] = "DPS";
  bossValue[1][18] = "Breakbar";
  bossValue[1][20] = "CondiCleans";
  bossValue[1][21] = "BoonStrips";

  for (i = 0; i < 6; i++) {
    bossSpecificSheet.getRange(2, i * 2 + 2, 1, 2).mergeAcross();
  }
  bossSpecificSheet.getRange(2, 15, 1, 2).mergeAcross();
  bossSpecificSheet.getRange(2, 17, 1, 2).mergeAcross();
  bossSpecificSheet.getRange(2, 19, 1, 2).mergeAcross();

  for (i = 0; i < 6; i++) {
    bossValue[2][i * 2 + 1] = "total";
    bossValue[2][i * 2 + 2] = "percent";
  }
  bossValue[2][13] = "AVG";
  bossValue[2][14] = "AVG";
  bossValue[2][15] = "lowest";
  bossValue[2][16] = "AVG";
  bossValue[2][17] = "highest";
  bossValue[2][18] = "AVG";
  bossValue[2][19] = "highest";
  bossValue[2][20] = "AVG";
  bossValue[2][21] = "AVG";

  // Participation total
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][1] =
      '=SUMPRODUCT(((Logs!C2:C=A1)+(Logs!C2:C=A1 & " CM"))*(MMULT((Logs!M2:V=A' +
      (i + 4) +
      ")*1;TRANSPOSE(COLUMN(Logs!M2:V)^0))>0))";
  }

  // Participation percent
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][2] = "=B" + (i + 4) + "/B31";
  }

  // Kills total
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][3] =
      '=SUMPRODUCT(((Logs!C2:C=A1)+(Logs!C2:C=A1 & " CM"))*(Logs!D2:D=TRUE)*(MMULT((Logs!M2:V=A' +
      (i + 4) +
      ")*1;TRANSPOSE(COLUMN(Logs!M2:V)^0))>0))";
  }

  // Kills percent
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][4] = "=D" + (i + 4) + "/B" + (i + 4);
  }

  // First Death total
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][5] =
      "=COUNTIFS(Logs!C2:C; A1; Logs!K2:K; A" +
      (i + 4) +
      ') + COUNTIFS(Logs!C2:C; A1 & " CM"; Logs!K2:K; A' +
      (i + 4) +
      ")";
  }

  // First Death percent
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][6] = "=F" + (i + 4) + "/B" + (i + 4);
  }

  // Downs total
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][7] =
      '=COUNTIFS(Logs!C2:C; A1;Logs!CE2:CE;"*" & A' +
      (i + 4) +
      '& "*";Logs!J2:J;FALSE) + COUNTIFS(Logs!C2:C; A1 & " CM";Logs!CE2:CE;"*" & A' +
      (i + 4) +
      '& "*";Logs!L2:L;FALSE)';
  }

  // Downs percent
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][8] = "=H" + (i + 4) + "/B" + (i + 4);
  }

  // Res total
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][9] =
      '=COUNTIFS(Logs!C2:C; A1; Logs!CH2:CH;"*" & A' +
      (i + 4) +
      '& "*") + COUNTIFS(Logs!C2:C; A1 & " CM"; Logs!CH2:CH;"*" & A' +
      (i + 4) +
      '& "*")';
  }

  // Res percent
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][10] = "=J" + (i + 4) + "/B" + (i + 4);
  }

  // Deads total
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][11] =
      '=COUNTIFS(Logs!C2:C; A1; Logs!CF2:CF;"*" & A' +
      (i + 4) +
      '& "*") + COUNTIFS(Logs!C2:C; A1 & " CM"; Logs!CF2:CF;"*" & A' +
      (i + 4) +
      '& "*")';
  }

  // Deads percent
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][12] = "=L" + (i + 4) + "/B" + (i + 4);
  }

  // ResTime
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][13] =
      '=SUMPRODUCT(((Logs!C2:C=A1)+(Logs!C2:C=A1 & " CM"))*(Logs!M2:V=A' +
      (i + 4) +
      ")*(Logs!BA2:BJ)) / J" +
      (i + 4);
  }

  // Damage Taken AVG
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][14] =
      '=SUMPRODUCT(((Logs!C2:C=A1)+(Logs!C2:C=A1 & " CM"))*(Logs!D2:D = TRUE)*(Logs!M2:V=A' +
      (i + 4) +
      ")*(Logs!AQ2:AZ)) / D" +
      (i + 4);
  }

  // Damage taken lowest
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][15] =
      '=MIN(MAP(FILTER(Logs!M2:V; (Logs!C2:C = A1)+(Logs!C2:C = A1 & " CM");Logs!D2:D = TRUE; (Logs!M2:M = A' +
      (i + 4) +
      ") + (Logs!N2:N = A" +
      (i + 4) +
      ") + (Logs!O2:O = A" +
      (i + 4) +
      ") + (Logs!P2:P = A" +
      (i + 4) +
      ") + (Logs!Q2:Q = A" +
      (i + 4) +
      ") + (Logs!R2:R = A" +
      (i + 4) +
      ") + (Logs!S2:S = A" +
      (i + 4) +
      ") + (Logs!T2:T = A" +
      (i + 4) +
      ") + (Logs!U2:U = A" +
      (i + 4) +
      ") + (Logs!V2:V = A" +
      (i + 4) +
      '));FILTER(Logs!AQ2:AZ; (Logs!C2:C = A1)+(Logs!C2:C = A1 & " CM");Logs!D2:D = TRUE; (Logs!M2:M = A' +
      (i + 4) +
      ") + (Logs!N2:N = A" +
      (i + 4) +
      ") + (Logs!O2:O = A" +
      (i + 4) +
      ") + (Logs!P2:P = A" +
      (i + 4) +
      ") + (Logs!Q2:Q = A" +
      (i + 4) +
      ") + (Logs!R2:R = A" +
      (i + 4) +
      ") + (Logs!S2:S = A" +
      (i + 4) +
      ") + (Logs!T2:T = A" +
      (i + 4) +
      ") + (Logs!U2:U = A" +
      (i + 4) +
      ") + (Logs!V2:V = A" +
      (i + 4) +
      "));LAMBDA(pers;value;IF(pers=A" +
      (i + 4) +
      "; value;999999))))";
  }

  // DPS AVG
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][16] =
      '=SUMPRODUCT(((Logs!C2:C=A1)+(Logs!C2:C=A1 & " CM"))*(Logs!D2:D = TRUE)*(Logs!M2:V=A' +
      (i + 4) +
      ")*(Logs!W2:AF)) / D" +
      (i + 4);
  }

  // DPS highest
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][17] =
      '=MAX(MAP(FILTER(Logs!M2:V; (Logs!C2:C = A1)+(Logs!C2:C = A1 & " CM");Logs!D2:D = TRUE; (Logs!M2:M = A' +
      (i + 4) +
      ") + (Logs!N2:N = A" +
      (i + 4) +
      ") + (Logs!O2:O = A" +
      (i + 4) +
      ") + (Logs!P2:P = A" +
      (i + 4) +
      ") + (Logs!Q2:Q = A" +
      (i + 4) +
      ") + (Logs!R2:R = A" +
      (i + 4) +
      ") + (Logs!S2:S = A" +
      (i + 4) +
      ") + (Logs!T2:T = A" +
      (i + 4) +
      ") + (Logs!U2:U = A" +
      (i + 4) +
      ") + (Logs!V2:V = A" +
      (i + 4) +
      '));FILTER(Logs!W2:AF; (Logs!C2:C = A1)+(Logs!C2:C = A1 & " CM");Logs!D2:D = TRUE; (Logs!M2:M = A' +
      (i + 4) +
      ") + (Logs!N2:N = A" +
      (i + 4) +
      ") + (Logs!O2:O = A" +
      (i + 4) +
      ") + (Logs!P2:P = A" +
      (i + 4) +
      ") + (Logs!Q2:Q = A" +
      (i + 4) +
      ") + (Logs!R2:R = A" +
      (i + 4) +
      ") + (Logs!S2:S = A" +
      (i + 4) +
      ") + (Logs!T2:T = A" +
      (i + 4) +
      ") + (Logs!U2:U = A" +
      (i + 4) +
      ") + (Logs!V2:V = A" +
      (i + 4) +
      "));LAMBDA(pers;value;IF(pers=A" +
      (i + 4) +
      "; value;0))))";
  }

  // Breakbar AVG
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][18] =
      '=SUMPRODUCT(((Logs!C2:C=A1)+(Logs!C2:C=A1 & " CM"))*(Logs!D2:D = TRUE)*(Logs!M2:V=A' +
      (i + 4) +
      ")*(Logs!AG2:AP)) / D" +
      (i + 4);
  }

  // Breakbar highest
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][19] =
      '=MAX(MAP(FILTER(Logs!M2:V; (Logs!C2:C = A1)+(Logs!C2:C = A1 & " CM");Logs!D2:D = TRUE; (Logs!M2:M = A' +
      (i + 4) +
      ") + (Logs!N2:N = A" +
      (i + 4) +
      ") + (Logs!O2:O = A" +
      (i + 4) +
      ") + (Logs!P2:P = A" +
      (i + 4) +
      ") + (Logs!Q2:Q = A" +
      (i + 4) +
      ") + (Logs!R2:R = A" +
      (i + 4) +
      ") + (Logs!S2:S = A" +
      (i + 4) +
      ") + (Logs!T2:T = A" +
      (i + 4) +
      ") + (Logs!U2:U = A" +
      (i + 4) +
      ") + (Logs!V2:V = A" +
      (i + 4) +
      '));FILTER(Logs!AG2:AP; (Logs!C2:C = A1)+(Logs!C2:C = A1 & " CM");Logs!D2:D = TRUE; (Logs!M2:M = A' +
      (i + 4) +
      ") + (Logs!N2:N = A" +
      (i + 4) +
      ") + (Logs!O2:O = A" +
      (i + 4) +
      ") + (Logs!P2:P = A" +
      (i + 4) +
      ") + (Logs!Q2:Q = A" +
      (i + 4) +
      ") + (Logs!R2:R = A" +
      (i + 4) +
      ") + (Logs!S2:S = A" +
      (i + 4) +
      ") + (Logs!T2:T = A" +
      (i + 4) +
      ") + (Logs!U2:U = A" +
      (i + 4) +
      ") + (Logs!V2:V = A" +
      (i + 4) +
      "));LAMBDA(pers;value;IF(pers=A" +
      (i + 4) +
      "; value;0))))";
  }

  // Condi Cleans
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][20] =
      '=SUMPRODUCT(((Logs!C2:C=A1)+(Logs!C2:C=A1 & " CM"))*(Logs!D2:D = TRUE)*(Logs!M2:V=A' +
      (i + 4) +
      ")*(Logs!BK2:BT)) / D" +
      (i + 4);
  }

  // BoonStrips
  for (i = 0; i < 10; i++) {
    bossValue[i + 3][21] =
      '=SUMPRODUCT(((Logs!C2:C=A1)+(Logs!C2:C=A1 & " CM"))*(Logs!D2:D = TRUE)*(Logs!M2:V=A' +
      (i + 4) +
      ")*(Logs!BU2:CD)) / D" +
      (i + 4);
  }

  if (bossMechanics.length > 0) {
    for (i = 0; i < bossMechanics.length; i++) {
      bossValue[15][i + 1] = bossMechanics[i];
      bossValue[16][i + 1] = "AVG";
      for (j = 0; j < 10; j++) {
        bossValue[j + 17][i + 1] =
          '=ARRAYFORMULA(COUNTIF(SPLIT(INDEX(Logs!A2:ZZ; 0; MATCH("' +
          bossMechanics[i] +
          '"; Logs!A1:ZZ1; 0)); ","); A' +
          (j + 18) +
          ")) / B" +
          (j + 4);
      }
    }
  }

  bossValue[28][0] = "AVG Duration";
  bossValue[28][1] =
    '=(SUMIFS(Logs!G2:G; Logs!C2:C; A1; Logs!D2:D; TRUE) + SUMIFS(Logs!G2:G; Logs!C2:C; A1 & " CM"; Logs!D2:D; TRUE)) / B32 / 86400000';
  bossValue[29][0] = "fastes Kill";
  bossValue[29][1] =
    '=HYPERLINK(INDEX(Logs!B2:B; MATCH(MIN(FILTER(Logs!G2:G; (Logs!C2:C = A1 & " CM")+(Logs!C2:C = A1); Logs!D2:D = TRUE; Logs!G2:G > 0)); Logs!G2:G; 0)); MIN(FILTER(Logs!G2:G; (Logs!C2:C = A1 & " CM")+(Logs!C2:C = A1); Logs!D2:D = TRUE; Logs!G2:G > 0)) / 86400000)';
  bossValue[30][0] = "amount of Tries";
  bossValue[30][1] = '=COUNTIF(Logs!C2:C; A1) + COUNTIF(Logs!C2:C; A1 & " CM")';
  bossValue[31][0] = "amount of Kills";
  bossValue[31][1] =
    '=COUNTIFS(Logs!C2:C; A1; Logs!D2:D; TRUE) + COUNTIFS(Logs!C2:C; A1 & " CM"; Logs!D2:D; TRUE)';
  bossValue[32][0] = "thereof in CM";
  bossValue[32][1] =
    '=COUNTIFS(Logs!C2:C; A1; Logs!D2:D; TRUE; Logs!J2:J; TRUE) + COUNTIFS(Logs!C2:C; A1 & " CM"; Logs!D2:D; TRUE; Logs!J2:J; TRUE)';
  bossValue[33][0] = "Kills with 0 Deaths";
  bossValue[33][1] =
    '=COUNTIFS(Logs!C2:C; A1; Logs!CF2:CF; FALSE) + COUNTIFS(Logs!C2:C; A1 & " CM"; Logs!CF2:CF; FALSE)';
  bossValue[34][0] = "Kills with 0 Downs";
  bossValue[34][1] =
    '=COUNTIFS(Logs!C2:C; A1; Logs!CE2:CE; FALSE) + COUNTIFS(Logs!C2:C; A1 & " CM"; Logs!CE2:CE; FALSE)';
  bossValue[35][0] = "SuccessRate";
  bossValue[35][1] = "=B32/B31";

  bossSpecificSheet.getRange(1, 1).setFontSize(14).setFontWeight("bold");
  bossSpecificSheet
    .getRange(2, 2, 2, 21)
    .setHorizontalAlignment("center")
    .setFontFamily("Arial")
    .setFontSize(11)
    .setFontWeight("bold")
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

  bossSpecificSheet.getRange(2, 2, 1, 21).setBackground(gray);
  bossSpecificSheet
    .getRange(4, 1, 10, 22)
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

  bossSpecificSheet.getRange(4, 2, 10, 21).setHorizontalAlignment("center");

  bossSpecificSheet
    .getRange(4, 1, 5, 22)
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

  for (i = 0; i < 7; i++) {
    bossSpecificSheet
      .getRange(4, 1 + i * 2, 10, 1)
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
  }

  for (i = 0; i < 4; i++) {
    bossSpecificSheet
      .getRange(4, 14 + i * 2, 10, 1)
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
    bossSpecificSheet
      .getRange(4, 21, 10, 1)
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
  }
  if (bossMechanics.length > 0) {
    bossSpecificSheet
      .getRange(16, 2, 2, bossMechanics.length)
      .setHorizontalAlignment("center")
      .setFontFamily("Arial")
      .setFontSize(11)
      .setFontWeight("bold")
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

    bossSpecificSheet
      .getRange(16, 2, 1, bossMechanics.length)
      .setBackground(gray);

    bossSpecificSheet
      .getRange(18, 1, 10, bossMechanics.length + 1)
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        null,
        "black",
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );

    bossSpecificSheet
      .getRange(18, 2, 10, bossMechanics.length)
      .setHorizontalAlignment("center");

    bossSpecificSheet
      .getRange(18, 1, 5, bossMechanics.length + 1)
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
  }

  bossRange.setValues(bossValue);
  bossSpecificSheet.getRange(29, 2, 2, 1).setNumberFormat('m"m" s"s" 000"ms"');
  bossSpecificSheet.getRange(36, 2, 1, 1).setNumberFormat("#0.00%");
  for (i = 0; i < 6; i++) {
    bossSpecificSheet.getRange(4, 3 + i * 2, 10, 1).setNumberFormat("#0.00%");
  }
  bossSpecificSheet.getRange(4, 12, 10, 1).setNumberFormat("#,##0");
  bossSpecificSheet.getRange(4, 14, 10, 1).setNumberFormat("#,##0.0");
  for (i = 0; i < 4; i++) {
    bossSpecificSheet.getRange(4, 15 + i * 2, 10, 1).setNumberFormat("#,##0.0");
  }
  for (i = 0; i < 3; i++) {
    bossSpecificSheet.getRange(4, 16 + i * 2, 10, 1).setNumberFormat("#,##0");
  }
  bossSpecificSheet.getRange(4, 22, 10, 1).setNumberFormat("#,##0.0");

  if (bossMechanics.length > 0) {
    bossSpecificSheet
      .getRange(18, 2, 10, bossMechanics.length)
      .setNumberFormat("#,##0.0#");
  }

  bossSpecificSheet.autoResizeColumn(1).setColumnWidths(2, 16, 90);

  var successRateRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue(
      green,
      SpreadsheetApp.InterpolationType.NUMBER,
      "1"
    )
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.NUMBER,
      "0,75"
    )
    .setGradientMinpointWithValue(
      red,
      SpreadsheetApp.InterpolationType.NUMBER,
      "0"
    )
    .setRanges([bossSpecificSheet.getRange(36, 2, 1, 1)])
    .build();
  rules.push(successRateRule);

  var participationRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([bossSpecificSheet.getRange(4, 3, 10, 1)])
    .build();
  rules.push(participationRule);

  var killRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([bossSpecificSheet.getRange(4, 5, 10, 1)])
    .build();
  rules.push(killRule);

  var firstDeathRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(red)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(green)
    .setRanges([bossSpecificSheet.getRange(4, 7, 10, 1)])
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
    .setRanges([bossSpecificSheet.getRange(4, 9, 10, 1)])
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
    .setRanges([bossSpecificSheet.getRange(4, 11, 10, 1)])
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
    .setRanges([bossSpecificSheet.getRange(4, 13, 10, 1)])
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
    .setRanges([bossSpecificSheet.getRange(4, 14, 10, 1)])
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
    .setRanges([bossSpecificSheet.getRange(4, 15, 10, 1)])
    .build();
  rules.push(dmgTakenRule);

  var dmgTakenLowestRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(red)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(green)
    .setRanges([bossSpecificSheet.getRange(4, 16, 10, 1)])
    .build();
  rules.push(dmgTakenLowestRule);

  var dpsRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([bossSpecificSheet.getRange(4, 17, 10, 1)])
    .build();
  rules.push(dpsRule);

  var dpsHighestRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([bossSpecificSheet.getRange(4, 18, 10, 1)])
    .build();
  rules.push(dpsHighestRule);

  var breakbarRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([bossSpecificSheet.getRange(4, 19, 10, 1)])
    .build();
  rules.push(breakbarRule);

  var breakbarHighestRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([bossSpecificSheet.getRange(4, 20, 10, 1)])
    .build();
  rules.push(breakbarHighestRule);

  var condiCleanseRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(green)
    .setGradientMidpointWithValue(
      yellow,
      SpreadsheetApp.InterpolationType.PERCENTILE,
      "50"
    )
    .setGradientMinpoint(red)
    .setRanges([bossSpecificSheet.getRange(4, 21, 10, 1)])
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
    .setRanges([bossSpecificSheet.getRange(4, 22, 10, 1)])
    .build();
  rules.push(boonStripsRule);

  if (bossMechanics.length > 0) {
    for (i = 0; i < bossMechanics.length; i++) {
      var mechanicRule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint(red)
        .setGradientMidpointWithValue(
          yellow,
          SpreadsheetApp.InterpolationType.PERCENTILE,
          "50"
        )
        .setGradientMinpoint(green)
        .setRanges([bossSpecificSheet.getRange(18, 2 + i, 10, 1)])
        .build();
      rules.push(mechanicRule);
    }
  }

  bossSpecificSheet.setConditionalFormatRules(rules);
}
