/** create basic layout for the full spreadsheet
 *
 */
function createFullLayout() {
  createStatisticsLayout();
  createSettingsLayout();
  createLogsLayout();
  createSetupLayout();
  createDowntimeLayout();
}

function createBossLayouts() {
  // Wing 1
  createBossSpecificLayout("Vale Guardian", vgMechanics);
  createBossSpecificLayout("Gorseval the Multifarious", gorseMechanics);
  createBossSpecificLayout("Sabetha the Saboteur", sabMechanics);
  // Wing 2
  createBossSpecificLayout("Slothasor", slothMechanics);
  createBossSpecificLayout("Matthias Gabrel", mattMechanics);
  // Wing 3
  createBossSpecificLayout("Keep Construct", kcMechanics);
  createBossSpecificLayout("Xera", xeraMechanics);
  // Wing 4
  createBossSpecificLayout("Cairn", cairnMechanics);
  createBossSpecificLayout("Mursaat Overseer", moMechanics);
  createBossSpecificLayout("Samarog", samMechanics);
  createBossSpecificLayout("Deimos", deiMechanics);
  // Wing 5
  createBossSpecificLayout("Soulless Horror", shMechanics);
  createBossSpecificLayout("Dhuum", dhuumMechanics);
  // Wing 6
  createBossSpecificLayout("Conjured Amalgamate", caMechanics);
  createBossSpecificLayout("Twin Largos", twinsMechanics);
  createBossSpecificLayout("Qadim", qadimMechanics);
  // Wing 7
  createBossSpecificLayout("Cardinal Adina", adinaMechanics);
  createBossSpecificLayout("Cardinal Sabir", sabirMechanics);
  createBossSpecificLayout("Qadim the Peerless", qpeerMechanics);
  // EoD Strikes
  createBossSpecificLayout("Aetherblade Hideout", trinMechanics);
  createBossSpecificLayout("Xunlai Jade Junkyard", ankkaMechanics);
  createBossSpecificLayout("Kaineng Overlook", koMechanics);
  createBossSpecificLayout("Old Lion's Court", olcMechanics);
  // SotO Strikes
  createBossSpecificLayout("Cosmic Observatory", coMechanics);
  createBossSpecificLayout("Temple of Febe", tofMechanics);
}
