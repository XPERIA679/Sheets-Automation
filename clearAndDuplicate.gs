function duplicateAndClearSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var mainSheet = ss.getSheetByName("Daily Report");
    if (!mainSheet) {
        Logger.log("Main sheet not found!");
        return;
    }

    var today = new Date();
    var dayofWeek = today.getDay();

    if (dayofWeek === 0 || dayofWeek === 6) {
      Logger.log("Today is a weekend. No Sheet will be generated");
      return;
    }

    today.setDate(today.getDate());
    var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy");

    var newSheet = mainSheet.copyTo(ss);
    newSheet.setName(formattedDate);

    var mainSheetIndex = mainSheet.getIndex();

    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(mainSheetIndex + 1);

    var rangeToClear = mainSheet.getRange("D12:G20");
    rangeToClear.clearContent();

    Logger.log("Sheet duplicated, renamed to " + formattedDate + ", and cells D12:G20 cleared.");
}