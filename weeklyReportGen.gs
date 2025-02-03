function generateAndSendWeeklyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName("Weekly Report");

  if (!reportSheet) {
    createHiddenReportSheet();
    reportSheet = ss.getSheetByName("Weekly Report");
  }

  reportSheet.clear();

  var today = new Date();
  if (today.getDay() !== 5) {
    Logger.log("Not Friday. Skipping report generation.");
    return; // Run only on Fridays
  }

  var startDate = new Date(today);
  startDate.setDate(today.getDate() - 4);

  var rawReport = "";

  for (var i = 0; i < 5; i++) {
    var dateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var dailySheet = ss.getSheetByName(dateStr);

    if (!dailySheet && i === 4) { 
      dailySheet = ss.getSheetByName("Daily Report");
    }

    if (dailySheet) {
      var data = dailySheet.getRange("D12:D20").getValues();
      rawReport += data.map(row => row.join("\t")).join("\n") + "\n";
    } else {
      Logger.log("No data found for " + (i === 4 ? "Daily Report (Friday)" : dateStr));
    }

    startDate.setDate(startDate.getDate() + 1);
  }

  if (!rawReport.trim()) {
    Logger.log("No weekly data found. Report not generated.");
    return;
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!apiKey) {
    Logger.log("❌ API_KEY is missing. Please set it in Script Properties.");
    return;
  }

  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;

  var prompt = "Format the following weekly report into a bulleted list of unique tasks only, with no time or status details:\n\n" + rawReport + "\n\nPlease remove any duplicate tasks.";

  var payload = JSON.stringify({
    "contents": [{ "parts": [{ "text": prompt }] }]
  });

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": payload
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var jsonResponse = JSON.parse(response.getContentText());

    if (!jsonResponse || !jsonResponse.candidates || jsonResponse.candidates.length === 0) {
      Logger.log("❌ Invalid response from Gemini API: " + response.getContentText());
      return;
    }

    var formattedReport = jsonResponse.candidates[0].content.parts[0].text;
    Logger.log("✅ Successfully formatted report.");

    // 🔹 Save formatted report in the hidden sheet
    reportSheet.getRange(1, 1).setValue("📊 Weekly Report");
    reportSheet.getRange(2, 1).setValue(formattedReport);

    // 🔹 Send to Google Chat
    var chatWebhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
    if (!chatWebhookUrl) {
      Logger.log("❌ CHAT_WEBHOOK_URL is missing. Please set it in Script Properties.");
      return;
    }

    var chatPayload = {
      "text": "📊 *Weekly Report* 📊\n" + formattedReport
    };

    var chatOptions = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(chatPayload)
    };

    var chatResponse = UrlFetchApp.fetch(chatWebhookUrl, chatOptions);
    Logger.log("✅ Google Chat response: " + chatResponse.getContentText());

  } catch (error) {
    Logger.log("❌ Error: " + error.toString());
  }
}
