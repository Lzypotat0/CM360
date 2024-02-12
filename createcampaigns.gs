/**
 * A helper function which creates one campaign via DCM API using information
 * from the sheet.
 * @param {object} ss Spreadsheet class object representing
 * current active spreadsheet
 * @param {Array} singleCampaignArray An array containing campaign information
 * @return {object} Campaign object
 */
function _createOneCampaign(ss, singleCampaignArray){
  var profileID = _fetchProfileId();

  var advertiserId = singleCampaignArray[0];
  var name = singleCampaignArray[1];
  var defaultLandingPageId = singleCampaignArray[2];
  var startDate = Utilities.formatDate(
      singleCampaignArray[3], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var endDate = Utilities.formatDate(
      singleCampaignArray[4], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');

  var campaignResource = {
    "kind": "dfareporting#campaign",
    "advertiserId": advertiserId,
    "name": name,
    "startDate": startDate,
    "endDate": endDate,
    "defaultLandingPageId":defaultLandingPageId
  };
  var newCampaign = DoubleClickCampaigns.Campaigns
      .insert(campaignResource, profileID);
  return newCampaign;
}



/**
 * Check if a campaign with the given ID exists.
 * Modify this function with your own implementation to interact with
 * the platform's API.
 * @param {string} campaignID The campaign ID to check for existence
 * @returns {boolean} True if the campaign exists, false otherwise
 */
function isCampaignExist(campaignID) {
  var profileID = _fetchProfileId();

  try {
    DoubleClickCampaigns.Campaigns.get(profileID, campaignID);
    return true;
  } catch (error) {
    return false;
  }
}


/**
 * Read campaign information from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */
function createCampaigns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CAMPAIGNS_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // exclude header row
    var campaignID = values[i][5]; // Assuming campaign ID is in column F (index 5)
    var newCampaign = null;

    if (!campaignID || !isCampaignExist(campaignID)) {
      newCampaign = _createOneCampaign(ss, values[i]);
    }

    var rowNum = i + 1;

    if (newCampaign) {
      sheet.getRange('F' + rowNum)
        .setValue(newCampaign.id)
        .setBackground('#b6d7a8');
    } else if (campaignID) {
      sheet.getRange('F' + rowNum)
        .setBackground('#ffc0cb');
    }
  }
  SpreadsheetApp.getUi().alert('Finished creating campaigns!');
}
