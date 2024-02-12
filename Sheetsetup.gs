

// Global variables/configurations
var DCMProfileID = 'DCMProfileID';
var AUTO_POP_HEADER_COLOR = '#a4c2f4';
var USER_INPUT_HEADER_COLOR = '#f4d8ba';
var AUTO_POP_CELL_COLOR = 'lightgray';
var Atomic_Colour = '#faa61a';
var Atomic_Colour2 = '#e45f3d';
var PASTEL_GREEN = '#b6d7a8';
var LIGHT_PINK = '#ffc0cb';

// Data range values
var DCMUserProfileID = 'DCMUserProfileID';

// sheet names
var SETUP_SHEET = 'Setup';
var SITES_SHEET = 'Sites';
var CAMPAIGNS_SHEET = 'Campaigns';
var PLACEMENTS_SHEET = 'Placements';
var ADS_SHEET = 'Ads';
var CREATIVES_SHEET = 'Creatives';
var LANDING_PAGES_SHEET = 'LandingPages';


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



