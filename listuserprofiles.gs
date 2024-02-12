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
