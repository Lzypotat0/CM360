/**
 * Using DCM API list all the sites this profile has added
 * and print them out on the sheet.
 */
function listSites() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SITES_SHEET);
  var profileID = _fetchProfileId();
  initializeSheet_(SITES_SHEET, true);

  // setup header row
  sheet.getRange('A1')
      .setValue('Site Name')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');
  sheet.getRange('B1')
      .setValue('Directory Site ID')
      .setFontWeight('bold')
      .setBackground('#a4c2f4');

  var sites = DoubleClickCampaigns.Sites.list(profileID).sites;
  for (var i = 0; i < sites.length; i++) {
    var currentObject = sites[i];
    var rowNum = i+2;
    sheet.getRange('A' + rowNum)
        .setValue(currentObject.name)
        .setBackground('lightgray');
    sheet.getRange('B' + rowNum)
        .setNumberFormat('@')
        .setValue(currentObject.directorySiteId)
        .setBackground('lightgray');
  }
}