/**
 * Read placement information from the sheet and use DCM API to bulk create them
 * in the DCM Account.
 */
function createPlacements() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PLACEMENTS_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // skip header row
    var placementID = values[i][9]; // Assuming placement ID is in column J (index 9)
    var newPlacement = null;

    if (!placementID || !isPlacementExist(placementID)) {
      newPlacement = _createOnePlacement(ss, values[i]);
    }

    var rowNum = i + 1;

    if (newPlacement) {
      sheet.getRange('J' + rowNum)
        .setValue(newPlacement.id)
        .setBackground('#b6d7a8')
    } else if (placementID) {
      sheet.getRange('J' + rowNum)
        .setBackground('#ffc0cb');
    }
  }
  SpreadsheetApp.getUi().alert('Congratulations! The robots have taken over.');
}



/**
 * A helper function which creates one placement via DCM API using information
 * from the sheet.
 * @param {object} ss Spreadsheet class object representing current active
 * spreadsheet
 * @param {Array} singlePlacementArray An array containing
 * placement information
 * @return {object|null} Placement object or null if already exists
 */
function _createOnePlacement(ss, singlePlacementArray) {
  var profileID = _fetchProfileId();

  var campaignID = singlePlacementArray[0];
  var name = singlePlacementArray[1];
  var siteId = singlePlacementArray[2];
  var paymentSource = 'PLACEMENT_AGENCY_PAID';
  var compatibility = (singlePlacementArray[3]).trim().toUpperCase();
  var size = singlePlacementArray[4];
  var sizeSplitted = size.split('x');

  var pricingScheduleStartDate = Utilities.formatDate(
    singlePlacementArray[5], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var pricingScheduleEndDate = Utilities.formatDate(
    singlePlacementArray[6], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var pricingSchedulePricingType = singlePlacementArray[7];
  var tagFormats = (singlePlacementArray[8]).split(',');
  for (var i = 0; i < tagFormats.length; i++) {
    tagFormats[i] = (tagFormats[i].trim()).replace(/\r?\n|\r/g, ', ');
  }

  var placementResource = {
    "kind": "dfareporting#placement",
    "campaignId": campaignID,
    "name": name,
    "directorySiteId": siteId,
    "paymentSource": paymentSource,
    "compatibility": compatibility,
    "size": {
      "width": sizeSplitted[0].trim(),
      "height": sizeSplitted[1].trim()
    },
    "pricingSchedule": {
      "startDate": pricingScheduleStartDate,
      "endDate": pricingScheduleEndDate,
      "pricingType": pricingSchedulePricingType
    },
    "tagFormats": tagFormats
  };

  var newPlacement = DoubleClickCampaigns.Placements
    .insert(placementResource, profileID);

  return newPlacement;
}

/**
 * Check if a placement with the given ID exists.
 * Modify this function with your own implementation to interact with
 * the platform's API.
 * @param {string} placementID The placement ID to check for existence
 * @returns {boolean} True if the placement exists, false otherwise
 */
function isPlacementExist(placementID) {
  var profileID = _fetchProfileId();

  try {
    DoubleClickCampaigns.Placements.get(profileID, placementID);
    return true;
  } catch (error) {
    return false;
  }
}
