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
