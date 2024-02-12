
/**
 * A helper function that retrieves a creative with matching dimensions
 * from the creatives sheet and matches the placement ID.
 * @param {object} ss Spreadsheet class object representing
 * current active spreadsheet
 * @param {object} size Object containing width and height of the creative
 * @param {string} placementId Placement ID to match against
 * @return {object|null} Creative object if a matching creative is found, null otherwise
 */
function _getCreativeWithMatchingDimensions(ss, size, placementId) {
  var sheet = ss.getSheetByName(CREATIVES_SHEET);
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) { // exclude header row
    var creativePlacementId = values[i][1]; // Column containing placement ID in the creatives sheet
    var creativeSize = {
      "width": values[i][2],
      "height": values[i][3]
    };

    if (creativeSize.width == size.width && creativeSize.height == size.height && creativePlacementId == placementId) {
      var creativeId = values[i][0]; // Column containing creative ID in the creatives sheet
      return { "id": creativeId };
    }
  }

  return null;
}



