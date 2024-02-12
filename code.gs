/**
 * Helper function to get DCM Profile ID.
 * @return {object} DCM Profile ID.
 */
function _fetchProfileId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName(DCMUserProfileID);
  return range.getValue();
}
/*******************************************************************************************************************
 * Find and clear, or create a new sheet named after the input argument.
 * @param {string} sheetName The name of the sheet which should be initialized.
 * @param {boolean} lock To lock the sheet after initialization or not
 * @return {object} A handle to a sheet.
 */
function initializeSheet_(sheetName, lock) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  if (lock) {
    sheet.protect().setWarningOnly(true);
  }
  return sheet;
}


/**
 * Initialize all tabs and their header rows
 */
function setupTabs() {
  _setupSetupSheet();
  _setupSitesSheet();
  _setupCampaignsSheet();
  _setupPlacementsGroupsSheet();
  _setupAdsSheet();
  _setupCreativesSheet();
  _setupLandingPagesSheet();
}


function _setupCampaignsSheet() {
}
/**
 * Initialize the Setup sheet and its header row
 * @return {object} A handle to the sheet.
*/

/**
 * Initialize the Sites sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupSitesSheet() {
  var sheet = initializeSheet_(SITES_SHEET, true);

  sheet.getRange('A1')
      .setValue('Site Name')
      .setBackground(AUTO_POP_HEADER_COLOR);
  sheet.getRange('B1')
      .setValue('Directory Site ID')
      .setBackground(AUTO_POP_HEADER_COLOR);
  sheet.getRange('A1:B1').setFontWeight('bold').setWrap(true);
  return sheet;
}

/*******************************************************************************************************************
 * Initialize the Campaigns sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupCampaignsSheet() {
  var sheet = initializeSheet_(CAMPAIGNS_SHEET, false);

  sheet.getRange('A1')
      .setValue('DCM Advertiser ID*')
      .setBackground(PASTEL_GREEN);
  sheet.getRange('B1')
      .setValue('Campaign Name*')
      .setBackground(PASTEL_GREEN);
  sheet.getRange('C1')
      .setValue('Landing Page ID*')
      .setBackground(PASTEL_GREEN);
  sheet.getRange('D1')
      .setValue('Start Date*')
      .setBackground(PASTEL_GREEN);
  sheet.getRange('E1')
      .setValue('End Date*')
      .setBackground(PASTEL_GREEN);
  sheet.getRange('F1')
      .setValue('Campaign ID (AUTO-POPULATED)')
      .setBackground(AUTO_POP_HEADER_COLOR);
  sheet.getRange('A1:F1')
      .setFontWeight('bold')
      .setWrap(true);
  return sheet;

}

/*******************************************************************************************************************
 * Initialize the Placements sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupPlacementsGroupsSheet() {
  var sheet = initializeSheet_(PLACEMENTS_SHEET, false);

  sheet.getRange('A1').setValue('Campaign ID*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('B1').setValue('Placement Name*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C1').setValue('Site ID*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('D1').setValue('Compatibility*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('E1').setValue('Size*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('F1').setValue('Pricing Schedule Start Date*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('G1').setValue('Pricing Schedule End Date*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('H1').setValue('Pricing Schedule Pricing Type*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('I1').setValue('Tag Formats*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('J1').setValue('Placement ID (do not edit; auto-filling)').setBackground(AUTO_POP_HEADER_COLOR);

  sheet.getRange('A1:J1').setFontWeight('bold').setWrap(true);

  // Set data validation for cell D2:D20
  var dataValidationRuleD = SpreadsheetApp.newDataValidation().requireValueInList(['DISPLAY', 'DISPLAY_INTERSTITIAL', 'IN_STREAM_AUDIO', 'IN_STREAM_VIDEO']).build();
  sheet.getRange('D2:D20').setDataValidation(dataValidationRuleD);

  // Set data validation for cell H2:H20
  var dataValidationRuleH = SpreadsheetApp.newDataValidation().requireValueInList(['PRICING_TYPE_CPA', 'PRICING_TYPE_CPC', 'PRICING_TYPE_CPM', 'PRICING_TYPE_CPM_ACTIVEVIEW', 'PRICING_TYPE_FLAT_RATE_CLICKS', 'PRICING_TYPE_FLAT_RATE_IMPRESSIONS']).build();
  sheet.getRange('H2:H20').setDataValidation(dataValidationRuleH);

  // Set data validation for cell I2:I20
  var dataValidationRuleI = SpreadsheetApp.newDataValidation().requireValueInList(['PLACEMENT_TAG_Iframe/JavaScript','PLACEMENT_TAG_IFRAME_JAVASCRIPT','PLACEMENT_TAG_INTERNAL_REDIRECT', 'PLACEMENT_TAG_JAVASCRIPT', 'PLACEMENT_TAG_STANDARD','PLACEMENT_TAG_CLICK_TRACKER','PLACEMENT_TAG_TRACKING','PLACEMENT_TAG_TRACKING_IFRAME','PLACEMENT_TAG_TRACKING_JAVASCRIPT']).build();
  sheet.getRange('I2:I20').setDataValidation(dataValidationRuleI);


  // Set data validation for cell E2:E20
  var dataValidationRuleE = SpreadsheetApp.newDataValidation().requireValueInList(['1x1', '88x31', '120x60', '120x90', '125x125', '120x240', '234x15', '234x60', '468x60', '320x50', '300x50', '320x100', '300x100', '320x150', '180x150', '240x400', '250x250', '200x200', '139x139', '300x200', '400x200', '468x600', '480x60', '336x280', '300x250', '580x400', '480x300', '360x300', '320x240', '240x320', '360x592', '375x667', '360x640', '530x442', '600x120', '320x320', '400x400', '600x150', '600x160', '600x200', '500x500', '600x250', '600x300', '768x90', '650x500', '600x400', '720x300', '800x100', '640x360', '600x500', '800x250', '970x66', '728x90', '800x400', '900x500', '900x600', '640x480', '800x600', '600x1200', '900x700', '768x1024', '1024x768', '800x800', '930x180', '980x90', '970x90', '980x120', '1000x90', '1000x200', '1000x250', '1000x260', '970x250', '980x240', '300x600', '1000x300', '1200x170', '300x1050', '1000x627', '1200x600', '1000x750', '1200x627', '1200x628', '1000x800', '1200x630', '1080x1080', '1200x675', '1200x800', '1200x900', '627x627', '1200x1200', '400x800', '1224x1584', '1980x1080']
).build();
  sheet.getRange('E2:E20').setDataValidation(dataValidationRuleE);
  return sheet;

}





/*******************************************************************************************************************
 * Initialize the Advertisers sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupAdsSheet() {
  var sheet = initializeSheet_(ADS_SHEET, false);

  sheet.getRange('A1').setValue('Campaign ID*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('B1').setValue('Ad Name*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C1').setValue('Start Date and Time*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('D1').setValue('End Date and Time*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('E1').setValue('Impression Ratio*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('F1').setValue('Priority*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('G1').setValue('Type*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('H1').setValue('Placement ID*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('I1').setValue('Ad ID (AUTO-POPULATED)').setBackground(AUTO_POP_HEADER_COLOR);

  sheet.getRange('A1:I1').setFontWeight('bold').setWrap(true);

  // Set data validation for cell F2:F20
  var dataValidationRuleF = SpreadsheetApp.newDataValidation().requireValueInList(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16']).build();
  sheet.getRange('F2:F20').setDataValidation(dataValidationRuleF);
  // Set data validation for cell E2:E20
  var dataValidationRuleE = SpreadsheetApp.newDataValidation().requireValueInList(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10']).build();
  sheet.getRange('E2:E20').setDataValidation(dataValidationRuleE);

  // Set data validation for cell G2:G20
  var dataValidationRuleG = SpreadsheetApp.newDataValidation().requireValueInList(['AD_SERVING_CLICK_TRACKER', 'AD_SERVING_DEFAULT_AD', 'AD_SERVING_STANDARD_AD', 'AD_SERVING_TRACKING']).build();
  sheet.getRange('G2:G20').setDataValidation(dataValidationRuleG);

  return sheet;
}



/*******************************************************************************************************************
 * Initialize the Creatives sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupCreativesSheet() {
  var sheet = initializeSheet_(CREATIVES_SHEET, false);

  sheet.getRange('A1').setValue('Advertiser ID*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('B1').setValue('Creative Name*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C1').setValue('Width*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('D1').setValue('Height*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('E1').setValue('Creative Type*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('F1').setValue('Creative Asset Type*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('G1').setValue('Creative Asset Name*').setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('H1').setValue('Creative ID (AUTO-POPULATED)').setBackground(AUTO_POP_HEADER_COLOR);

  sheet.getRange('A1:H1').setFontWeight('bold').setWrap(true);

  // Set data validation for cell E2:E20
  var dataValidationRuleE = SpreadsheetApp.newDataValidation().requireValueInList(['DISPLAY','RICH_MEDIA_DISPLAY_BANNER','IN-STREAM_VIDEO','IN-STREAM_VIDEO_REDIRECT','IN-STREAM_AUDIO','AUDIO_REDIRECT','CUSTOM_DISPLAY','CUSTOM_DISPLAY_INTERSTITIAL','DISPLAY_REDIRECT','TRACKING']).build();
  sheet.getRange('E2:E20').setDataValidation(dataValidationRuleE);

  // Set data validation for cell F2:F20
  var dataValidationRuleF = SpreadsheetApp.newDataValidation().requireValueInList(['AUDIO', 'FLASH', 'HTML', 'HTML_IMAGE', 'IMAGE', 'VIDEO']).build();
  sheet.getRange('F2:F20').setDataValidation(dataValidationRuleF);

  return sheet;
}



/*******************************************************************************************************************
 * Initialize the LandingPages sheet and its header row
 * @return {object} A handle to the sheet.
 */
function _setupLandingPagesSheet() {
  var sheet = initializeSheet_(LANDING_PAGES_SHEET, false);

  sheet.getRange('A1')
      .setValue("Advertiser ID*")
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('B1')
      .setValue("Landing Page Name*")
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('C1')
      .setValue("Landing Page URL*")
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('D1')
      .setValue("Landing Page ID (Do Not Edit; AUTOPOPULATED)")
      .setBackground(AUTO_POP_HEADER_COLOR);
      
  sheet.getRange("A1:H1").setFontWeight("bold").setWrap(true);
  return sheet;

  
}

function listUserProfiles() {
  try {
    const profiles = DoubleClickCampaigns.UserProfiles.list();

    if (profiles.items) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName("Setup");

      // Clear existing data range
      const dataRange = sheet.getRange("A9:D");
      dataRange.clearContent();

      // Write header row
      sheet.getRange(9, 1, 1, 4).setValues([["Advertiser ID","Advertiser","User Profile ID", "Username"]]);

      sheet.getRange('A9:D9')
        .setFontWeight('bold')
        .setWrap(true)
        .setBackground(Atomic_Colour)
        .setFontSize(14)
        .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

      const data = [];
      for (let i = 0; i < profiles.items.length; i++) {
        const profile = profiles.items[i];
        data.push([profile.accountId, profile.accountName, profile.profileId, profile.userName]);
      }

      // Write profile data starting from row 10
      if (data.length > 0) {
        sheet.getRange(10, 1, data.length, 4).setValues(data);
      }

      // Set data validation for cell C10:C30
      const validationRange = sheet.getRange("C10:C30");
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      sheet.getRange("C5")
        .setDataValidation(rule)
        .setBackground(LIGHT_PINK)
        .setFontSize(20);
    }
  } catch (e) {
    console.log('Failed with error: %s', e.error);
  }
}

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

function _setupSetupSheet() {  
  var sheet = initializeSheet_(SETUP_SHEET, false);
  var cell;
  
  sheet.getRange('A2:A2').setValue("DCM Bulk Trafficking");
  sheet.getRange('A2:C2')
      .setFontWeight('bold')
      .setWrap(true)
      .setBackground(AUTO_POP_HEADER_COLOR)
      .setFontSize(24)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

  sheet.getRange('A3')
      .setValue('For any questions contact Jameson - j.tang@atomic212.com.au');
  
  var instructions = [
     null,
     null,
     null,
     null,
     null,
     null,
     null,
    "Initial setup:",
    "# Make a copy of this template",
    "# Enter DCM Profile ID in C5 of this tab - [Data] > "+
    " [Named Ranges] > 'DCMUserProfileID', 'Setup!C5'",
    "# [DCM Functions] > [Fetch User Profile ID]",
    null,
    null,
    null,
    null,
    "How to use:",
    "# Select Profile ID in the options provided in cell C5",
    "# [Sites tab] Retrieve the list of sites and IDs by [DCM Functions] > [List Sites]",
    "# [Campaigns tab] Bulk create Campaigns by [DCM Functions] > [Bulk Create Campaigns]",
    "# [Placements tab]  Bulk create Placements groups by [DCM Functions] > [Bulk Create Placements]",
    "# [Ads tab] Bulk create Ads by [DCM Functions] > [Bulk Create Ads]",
    "# [Creatives tab] Bulk create Creatives by [DCM Functions] > [Bulk Create Creatives]",
    "# [LandingPages tab] Bulk create Landing Pages by [DCM Functions] > [Bulk Create Landing Pages]"
  ]
  
  for(var i=0; i<instructions.length; i++) {
    cell = i+2
    var count = instructions[i] == null ? -1 : (i==0 ? 0 : count+1);
    var value = instructions[i] == null ? null : instructions[i].replace('#', count + ')');
    sheet.getRange('F' + cell).setValue(value);
    
    if (count == 0) {
      sheet.getRange('F' + cell + ':K' + cell)
        .setFontWeight("bold")
        .setWrap(true)
        .setBackground(AUTO_POP_HEADER_COLOR)
        .setFontSize(12);
    }
  }
  sheet.getRange('F9:K15').setBorder(
    true,  // top
    true,  // left
    true,  // bottom
    true,  // right
    false, // vertical
    false, // horizontal
  "black",
  SpreadsheetApp.BorderStyle.THICK
);
  sheet.getRange('F9:K9').setBorder(
    true,  // top
    true,  // left
    true,  // bottom
    true,  // right
    false, // vertical
    false, // horizontal
  "black",
  SpreadsheetApp.BorderStyle.THICK
);

  sheet.getRange('F17:K24').setBorder(
    true,  // top
    true,  // left
    true,  // bottom
    true,  // right
    false, // vertical
    false, // horizontal
  "black",
  SpreadsheetApp.BorderStyle.THICK
);
  sheet.getRange('F17:K17').setBorder(
    true,  // top
    true,  // left
    true,  // bottom
    true,  // right
    false, // vertical
    false, // horizontal
  "black",
  SpreadsheetApp.BorderStyle.THICK
);

  sheet.getRange('F' + (cell+3)).setValue("Legend")
      .setFontWeight("bold")
      .setFontSize(16);
  sheet.getRange('F' + (cell+4))
      .setValue("Green Cells / Columns are for input");
  sheet.getRange('F' + (cell+5))
      .setValue("Blue Cells /AUTOPOPULATED (DO NOT EDIT)");
  
  sheet.getRange('F' + (cell+3))
      .setBackground(Atomic_Colour);
  sheet.getRange('F' + (cell+4) + ':H' + (cell+4))
      .setBackground(USER_INPUT_HEADER_COLOR);
  sheet.getRange('F' + (cell+5) + ':H' + (cell+5))
      .setBackground(AUTO_POP_HEADER_COLOR);
  
  sheet.getRange('B5').setValue("User Profile ID")
                      .setBackground(Atomic_Colour2)
                      .setFontSize(20);
  sheet.getRange('C5').setBackground(PASTEL_GREEN
  );

  sheet.getRange("B5:C5").setFontWeight("bold").setWrap(true);
  return sheet;
}


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

function listActiveCampaigns() {
  const profileId = '8091130'; // Replace with your profile ID.
  const fields = 'nextPageToken,campaigns(id,name)';
  const sheetName = 'Active Campaigns'; // Specify the name of the destination sheet.

  let result;
  let pageToken;
  let sheet;

  try {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      // If the sheet doesn't exist, create it.
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      sheet.appendRow(['ID', 'Name']); // Add column headers.
    } else {
      // Clear the contents of the sheet.
      sheet.clear();
      sheet.appendRow(['ID', 'Name']); // Add column headers.
    }

    do {
      result = DoubleClickCampaigns.Campaigns.list(profileId, {
        'archived': false,
        'fields': fields,
        'pageToken': pageToken
      });

      if (result.campaigns) {
        for (let i = 0; i < result.campaigns.length; i++) {
          const campaign = result.campaigns[i];
          sheet.appendRow([campaign.id, campaign.name]); // Append campaign data to the sheet.
        }
      }

      pageToken = result.nextPageToken;
    } while (pageToken);

    // Display a success message after the operation.
    SpreadsheetApp.getUi().alert('Active campaigns listed successfully.');

  } catch (e) {
    // Log the complete error object for troubleshooting.
    Logger.log('Failed with error:', e.message); // Log the error message for more details.
    // You can also log additional information if needed, e.g., e.stack for the stack trace.
  }
}


