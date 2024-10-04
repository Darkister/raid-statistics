/** Write Data of the Log into the Spreadsheet
 *  @param {Array} logs
 */
function writeDataIntoSpreadsheet(logs) {
  var row = logSheet.getLastRow() + 1,
    // columns = logSheet.getLastColumn(),
    statusCell = settingsSheet.getRange(13, 3),
    values = new Array();

  for (var i = 0; i < logs.length; i++) {
    var bossTargetsPosition = null;
    if (!(i in values)) {
      values.push([]);
    }
    try {
      var log = logs[i];
      statusCell.setValue("Calculating Logs " + i + "/" + logs.length);
      Logger.log("Next Log to calculate: " + log);
      var json = apiFetch(log);

      for (t = 0; t < json.targets.length; t++) {
        if (json.fightName.includes(json.targets[t].name)) {
          bossTargetsPosition = t;
          break;
        }
      }

      // Date
      values[i].push(getDayOfLog(json.timeStart));
      // Log
      values[i].push(log);
      // Boss/Encounter
      values[i].push(json.fightName);
      // Success
      values[i].push(json.success);
      // Rest HP
      values[i].push((100 - json.targets[0].healthPercentBurned) / 100);
      // Duration
      values[i].push(json.duration);
      // DurationMS
      values[i].push(json.durationMS);
      // timeStart
      values[i].push(json.timeStart.split(" ")[1]);
      // timeEnd
      values[i].push(json.timeEnd.split(" ")[1]);
      // CM?
      values[i].push(json.isCM);
      // First Death
      var firstDead = firstDeath(json);
      values[i].push(firstDead);
      // all Player down on First Death
      values[i].push(
        firstDead == false ? false : allPlayerDownOnFirstDeath(json)
      );
      // Players Accountname
      var players = getPlayer(json);
      for (p = 0; p < 10; p++) {
        values[i].push(players[p]);
      }
      // FullFight DPS
      var dps = getPlayersDPS(json, players, bossTargetsPosition);
      for (d = 0; d < 10; d++) {
        values[i].push(dps[d]);
      }
      // Breakbar
      for (p = 0; p < 10; p++) {
        values[i].push(json.players.filter(player =>
          player.account === players[p]).reduce((sum, player) =>
            sum + player.dpsTargets[0][0].breakbarDamage, 0
          )
        )
      }
      // Received Damage
      for (p = 0; p < 10; p++) {
        values[i].push(json.players.filter(player =>
          player.account === players[p]).reduce((sum, player) =>
            sum + player.defenses[0].damageTaken, 0
          )
        )
      }
      // Res Duration
      for (p = 0; p < 10; p++) {
        values[i].push(json.players.filter(player =>
          player.account === players[p]).reduce((sum, player) =>
            sum + player.support[0].resurrectTime, 0
          )
        )
      }
      // Condi Cleansed
      for (p = 0; p < 10; p++) {
        values[i].push(json.players.filter(player =>
          player.account === players[p]).reduce((sum, player) =>
            sum + player.support[0].condiCleanse + player.support[0].condiCleanseSelf, 0
          )
        )
      }
      // Boon Strips
      for (p = 0; p < 10; p++) {
        values[i].push(json.players.filter(player =>
          player.account === players[p]).reduce((sum, player) =>
            sum + player.support[0].boonStrips, 0
          )
        )
      }
      // Basic + Boss Mechanics
      var bossFilter = getBossFilter(json.fightName),
        combinedMechanics = combineMechanicsData(
          json.mechanics.filter((mechanic) =>
            bossFilter.includes(mechanic.fullName || mechanic.name)
          )
        );
      for (m = 0; m < combinedMechanics.length; m++) {
        var index = findOrCreateColumn(combinedMechanics[m]);
        while (values[i].length < index) {
          values[i].push("false");
        }
        values[i][index] = playersFailedMechanic(json, combinedMechanics[m]);
      }
    } catch (e) {
      console.error("apiFetch yielded error: " + e);
      Logger.log("Skip log");
    }
  }
  Logger.log(values);
  var max = Math.max(...values.map((i) => i.length));
  for (var value of values) {
    while (value.length < max) {
      value.push("false");
    }
  }
  var valuesRange = logSheet.getRange(row, 1, logs.length, max);
  valuesRange
    .setValues(values)
    .setBorder(
      null,
      true,
      null,
      true,
      true,
      null,
      "black",
      SpreadsheetApp.BorderStyle.SOLID
    )
    .setFontWeight("normal")
    .setFontFamily("Arial")
    .setFontSize("10")
    .setBackground("#FFFFFF")
    .setFontColor("#000000");
  logSheet.getRange(row, 1, logs.length, 1).setFontWeight("bold");
  logSheet.getRange(row, 2, logs.length, 1).setFontColor("#00B2EE");
  logSheet.getRange(row, 5, logs.length, 1).setNumberFormat("#0.00%");
}

/** Get data of a log as json
 *  @param {String} link  permalink of the Encounter
 *  @return {String}      returns the full encounterinformation as json
 */
function apiFetch(permalink) {
  var fetchUrl = "https://dps.report/getJson?permalink=",
    requestUrl = fetchUrl + permalink;
  var opt = {
    contentType: "application/json",
    muteHttpExceptions: true,
  },
    data = UrlFetchApp.fetch(requestUrl, opt);

  data = data.getContentText();
  return JSON.parse(data);
}

/** Get Accountname of first death player for given Encounter
 *  @param {String} json          fightData as json of the Encounter
 *  @return {String || boolean}   returns the first death player of the given fight or false if nobody died
 */
function firstDeath(json) {
  var mechanics = json.mechanics,
    players = json.players,
    deads = mechanics.find((mechanic) => mechanic.name === "Dead");

  Logger.log(deads);
  // if nobody died return false
  if (!deads) {
    return false;
  }

  var playername = deads.mechanicsData[0].actor,
    deadPlayer = players.find((player) => player.name === playername);

  if (deadPlayer) {
    return deadPlayer.account;
  } else {
    return false;
  }
}

/** Checks players of the given fight
 *  @param {String} json  fightData as json of the Encounter
 *  @return {String[]}    returns an Array which contains all players Accountnames
 */
function getPlayer(json) {
  var allPlayersInfo = json.players,
    players = new Array();

  for (var i = 0; i < allPlayersInfo.length; i++) {
    if (!players.includes(allPlayersInfo[i].account) && !allPlayersInfo[i].friendlyNPC) {
      players.push(allPlayersInfo[i].account);
    }
  }

  while (players.length < 10) {
    players.push("");
  }

  return players;
}

/** Get Players DPS of the given fight
 *  @param {String} json      fightData as json of the Encounter
 *  @param {String[]} players Array which contains all players Accountnames
 *  @param {Int} bossPos      Position of the BossTarget in data
 *  @return {Int[]}           returns array which contains DPS of players
 */
function getPlayersDPS(json, players, bossPos) {
  var dps = new Array();
  for (p = 0; p < players.length; p++) {
    if (players[p] == "") {
      dps.push(0);
    }
    else {
      var relevantPlayerData = json.players.filter(player =>
        player.account === players[p]
      )
      if (bossPos == null) {
        if (json.fightName.includes("Twin Largos")) {
          dps.push(relevantPlayerData.reduce((sum, player) =>
            sum + player.dpsTargets[0][0].dps + player.dpsTargets[1][0].dps, 0
          ));
        } else if (json.fightName.includes("Aetherblade Hideout")) {
          dps.push(relevantPlayerData.reduce((sum, player) =>
            sum + player.dpsTargets[0][0].dps + player.dpsTargets[3][0].dps, 0
          ));
        } else {
          dps.push(relevantPlayerData.reduce((sum, player) =>
            sum + player.dpsTargets[0][0].dps, 0
          ));
        }
      } else {
        dps.push(relevantPlayerData.reduce((sum, player) =>
          sum + player.dpsTargets[bossPos][0].dps, 0
        ));
      }
    }
  }

  while (dps.length < 10) {
    dps.push(0);
  }

  return dps;
}

/** Get the Day where the try was made
 *  @param {String} json  fightData as json of the Encounter
 *  @return {String}      returns a date
 */
function getDayOfLog(timeStart) {
  var date = timeStart.split("-"),
    year = date[0],
    month = date[1],
    day = date[2].split(" ")[0];
  return day + "." + month + "." + year;
}

/**
 *
 */
function allPlayerDownOnFirstDeath(json) {
  var mechanics = json.mechanics,
    dead,
    downs,
    res,
    firstDeathTime,
    lastDownTime;

  try {
    for (var i = 0; i < mechanics.length; i++) {
      if (mechanics[i].name == "Dead") {
        dead = mechanics[i].mechanicsData;
        break;
      }
    }

    for (var i = 0; i < mechanics.length; i++) {
      if (mechanics[i].name == "Downed") {
        downs = mechanics[i].mechanicsData;
        break;
      }
    }

    firstDeathTime = dead[0].time;
    lastDownTime = downs[downs.length - 1].time;
  } catch {
    return false;
  }
  try {
    for (var i = 0; i < mechanics.length; i++) {
      if (mechanics[i].name == "Got up") {
        res = mechanics[i].mechanicsData;
        break;
      }
    }

    var firstDownTime = downs[downs.length - 10].time;
    var lastResTime = res[res.length - 1].time;

    return lastResTime < firstDownTime && firstDeathTime > lastDownTime;
  } catch {
    return downs.length >= 10 && firstDeathTime > lastDownTime;
  }
}

/** Checks the Encounter and returns boss specific filter
 *  @param {String} fightName Name of the Encounter
 *  @return {String[]}        the boss specific Filter
 */
function getBossFilter(fightName) {
  if (fightName === "Tal-Wächter" || fightName === "Vale Guardian") {
    return basicMechanics.concat(vgMechanics);
  }
  if (fightName === "Gorseval" || fightName === "Gorseval the Multifarious") {
    return basicMechanics.concat(gorseMechanics);
  }
  if (fightName === "Sabetha" || fightName === "Sabetha the Saboteur") {
    return basicMechanics.concat(sabMechanics);
  }
  if (fightName === "Faultierion" || fightName === "Slothasor") {
    return basicMechanics.concat(slothMechanics);
  }
  if (fightName === "Matthias" || fightName === "Matthias Gabrel") {
    return basicMechanics.concat(mattMechanics);
  }
  if (
    fightName === "Festenkonstrukt" ||
    fightName === "Festenkonstrukt CM" ||
    fightName === "Keep Construct" ||
    fightName === "Keep Construct CM"
  ) {
    return basicMechanics.concat(kcMechanics);
  }
  if (fightName === "Xera" || fightName === "Xera") {
    return basicMechanics.concat(xeraMechanics);
  }
  if (fightName === "Cairn" || fightName === "Cairn CM") {
    return basicMechanics.concat(cairnMechanics);
  }
  if (fightName === "Mursaat Overseer" || fightName === "Mursaat Overseer CM") {
    return basicMechanics.concat(moMechanics);
  }
  if (fightName === "Samarog" || fightName === "Samarog CM") {
    return basicMechanics.concat(samMechanics);
  }
  if (fightName === "Deimos" || fightName === "Deimos CM") {
    return basicMechanics.concat(deiMechanics);
  }
  if (fightName === "Soulless Horror" || fightName === "Soulless Horror CM") {
    return basicMechanics.concat(shMechanics);
  }
  if (fightName === "Fluss der Seelen" || fightName === "River of Souls") {
    return basicMechanics.concat(rrMechanics);
  }
  if (fightName === "Statue of Ice") {
    return basicMechanics.concat(bkMechanics);
  }
  if (fightName === "Dhuum" || fightName === "Dhuum CM") {
    return basicMechanics.concat(dhuumMechanics);
  }
  if (
    fightName === "Conjured Amalgamate" ||
    fightName === "Conjured Amalgamate CM"
  ) {
    return basicMechanics.concat(caMechanics);
  }
  if (fightName === "Twin Largos" || fightName === "Twin Largos CM") {
    return basicMechanics.concat(twinsMechanics);
  }
  if (fightName === "Qadim" || fightName === "Qadim CM") {
    return basicMechanics.concat(qadimMechanics);
  }
  if (fightName === "Cardinal Adina" || fightName === "Cardinal Adina CM") {
    return basicMechanics.concat(adinaMechanics);
  }
  if (fightName === "Cardinal Sabir" || fightName === "Cardinal Sabir CM") {
    return basicMechanics.concat(sabirMechanics);
  }
  if (
    fightName === "Qadim the Peerless" ||
    fightName === "Qadim the Peerless CM"
  ) {
    return basicMechanics.concat(qpeerMechanics);
  }
  if (
    fightName === "Aetherblade Hideout" ||
    fightName === "Aetherblade Hideout CM"
  ) {
    return basicMechanics.concat(trinMechanics);
  }
  if (fightName === "Xunlai Jade Junkyard" || fightName === "Xunlai Jade Junkyard CM") {
    return basicMechanics.concat(ankkaMechanics);
  }
  if (fightName === "Kaineng Overlook" || fightName === "Kaineng Overlook CM") {
    return basicMechanics.concat(koMechanics);
  }
  if (fightName === "Old Lion's Court" || fightName === "Old Lion's Court CM") {
    return basicMechanics.concat(olcMechanics);
  }
  if (fightName === "Cosmic Observatory" || fightName === "Cosmic Observatory CM") {
    return basicMechanics.concat(coMechanics);
  }
  if (fightName === "Temple of Febe" || fightName === "Temple of Febe CM") {
    return basicMechanics.concat(tofMechanics);
  } else {
    return basicMechanics;
  }
}

/**
 *
 * @param {*} mechanics
 * @returns
 */
function combineMechanicsData(mechanics) {
  const combinedMechanics = [];
  const seenFullNames = new Set();

  mechanics.forEach((mechanic) => {
    // If mechanic.fullName doesn't exist, set it to mechanic.name
    const fullName =
      mechanic.fullName != null ? mechanic.fullName : mechanic.name;
    if (!seenFullNames.has(fullName)) {
      const combinedMechanic = {
        ...mechanic,
        mechanicsData: mechanics
          .filter(
            (m) => fullName === (m.fullName != null ? m.fullName : m.name)
          )
          .flatMap((m) => m.mechanicsData),
      };
      seenFullNames.add(fullName);
      combinedMechanics.push(combinedMechanic);
    }
  });

  return combinedMechanics;
}

/** Get info that players failed a given mechanic
 *  @param {String} json      fightData as json of the Encounter
 *  @param {String} mechanic  the Object of the mechanic to calculate
 *  @return {String}          returns a string with all players in a row who failed the given mechanic
 */
function playersFailedMechanic(json, mechanic) {
  return mechanic.mechanicsData
    .map(
      (mechs) =>
        json.players.find((player) => player.name === mechs.actor).account
    )
    .join(",");
}

/**
 * Find or create a Column by Name with a given Mechanic
 * @param {*} mechanic  the mechanic to search
 * @returns             the Index of the Column
 */
function findOrCreateColumn(mechanic) {
  var logHeaders = logSheet
    .getRange(1, 1, 1, logSheet.getLastColumn())
    .getValues()[0],
    mechanicName = mechanic.fullName || mechanic.name,
    columnIndex = logHeaders.indexOf(mechanicName);

  if (columnIndex === -1) {
    // If the title is not found, add a new column at the end
    columnIndex = logHeaders.length;
    logSheet.insertColumnAfter(logHeaders.length);
    logSheet
      .getRange(1, columnIndex + 1)
      .setValue(mechanicName)
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
  }

  // Return columnIndex with 0-based index
  return columnIndex;
}
