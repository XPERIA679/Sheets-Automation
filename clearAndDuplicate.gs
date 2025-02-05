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

    let reporterToday = newSheet.getRange("B3").getValue();
    newSheet.getRange("B3").setValue(reporterToday);

    let scratchSheet = ss.getSheetByName("Weekly Report");
    if (!scratchSheet) {
      Logger.log("Weekly Report sheet not found!");
      return;
    }

    let counter = scratchSheet.getRange("B4").getValue();
    if (typeof counter === 'number') {
        let incrementedCounter = counter + 1;
        scratchSheet.getRange("B4").setValue(incrementedCounter);
        let reporterTomorrow = mainSheet.getRange("B3").getValue();
        Logger.log("The reporter tomorrow is: " + reporterTomorrow);
    } else {
        Logger.log("Counter is not a number!");
    }   
}
