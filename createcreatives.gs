/**
 * Read creatives information from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */
function createCreatives() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CREATIVES_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // exclude header row
    var newCreative = _createOneCreative(ss, values[i]);
    var rowNum = i+1;
    sheet.getRange('H' + rowNum)
        .setValue(newCreative.id)
        .setBackground('lightgray');
  }

  SpreadsheetApp.getUi().alert('Finished creating the creatives!');
}

/**
 * Read landing pages information from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */


/**
 * A helper function which creates one creative via DCM API using information
 * from the sheet.
 * @param {object} ss Spreadsheet class object representing
 * current active spreadsheet
 * @param {Array} singleCreativeArray An array containing creative information
 * @return {object} Creative object
 */
function _createOneCreative(ss, singleCreativeArray){
  var profileID = _fetchProfileId();

  var advertiserId = singleCreativeArray[0];
  var name = singleCreativeArray[1];
  var width = singleCreativeArray[2];
  var height = singleCreativeArray[3];
  var creativeType = singleCreativeArray[4];
  var assetType = singleCreativeArray[5];
  var assetName = singleCreativeArray[6];

  var creativeResource =  {
    "name": name,
    "advertiserId": advertiserId,
    "size": {
      "width": width,
      "height": height
    },
    "active": true,
    "type": creativeType,
    "creativeAssets": [
      {
        "assetIdentifier": {
          "type": assetType,
          "name": assetName
        }
      }
    ]
  };

  var newCreative = DoubleClickCampaigns.Creatives
      .insert(creativeResource, profileID);
  return newCreative;

}