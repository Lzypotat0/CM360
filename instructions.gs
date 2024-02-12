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

