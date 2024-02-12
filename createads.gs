function createAds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ADS_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // exclude header row
    var newAd = _createOneAd(ss, values[i]);
    var rowNum = i + 1;
    if (newAd) {
      sheet.getRange('I' + rowNum)
        .setValue(newAd.id)
        .setBackground('#b6d7a8');
    } else {
      sheet.getRange('I' + rowNum)
        .setBackground('#ffc0cb');
    }
  }

  SpreadsheetApp.getUi().alert('Finished creating the ads!');
}



function _createOneAd(ss, singleAdArray) {
  var profileID = _fetchProfileId();

  if (!Array.isArray(singleAdArray) || singleAdArray.length < 9) {
    // Invalid or incomplete ad information, skip creating the ad
    return null;
  }

  var campaignId = singleAdArray[0];
  var name = singleAdArray[1];

  // Check if AD ID already exists
  var adId = singleAdArray[8]; // Assuming AD ID is in column 'I'
  if (adId && adExists(adId, profileID)) {
    // AD ID already exists, skip creating the ad
    return null;
  }

  var startTime = Utilities.formatDate(
    singleAdArray[2], ss.getSpreadsheetTimeZone(),
    'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');

  var endTime = Utilities.formatDate(
    singleAdArray[3], ss.getSpreadsheetTimeZone(),
    'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');

  var impressionRatio = singleAdArray[4];
  var priority = singleAdArray[5];
  var type = singleAdArray[6];
  var placementId = singleAdArray[7];

  // Construct the adResource object
  var adResource = {
    "kind": "dfareporting#ad",
    "campaignId": campaignId,
    "name": name,
    "startTime": startTime,
    "endTime": endTime,
    "deliverySchedule": {
      "impressionRatio": impressionRatio,
      "priority": "AD_PRIORITY_" + priority
    },
    "type": type,
    "placementAssignments": [
      {
        "placementId": placementId
      }
    ]
  };

  // Rest of the code for creating the ad...

  var newAd = DoubleClickCampaigns.Ads.insert(adResource, profileID);

  // Check if newAd is null before accessing its properties
  if (newAd && newAd.id) {
    return newAd;
  } else {
    return null;
  }
}

// The adExists function remains the same as before
function adExists(adId, profileId) {
  var ad = DoubleClickCampaigns.Ads.get(profileId, adId);
  return !!ad;
}
