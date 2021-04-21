function SendEmails() {

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails").activate();
  var ss=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr= ss.getLastRow();

  var EmailMessage=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange("B3").getDisplayValue();
  for (var i=2;i<=lr;i++){


    var Subject=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange("B2").getValue();
    var CurrentEmail=ss.getRange(i,1).getValue();
    var CurrentName=ss.getRange(i,2).getValue();
    var CurrentFile=ss.getRange(i,3).getValue();
    var CurrentMonths=ss.getRange(i,4).getValue();

    var messagebody=EmailMessage.replace("{Name}",CurrentName).replace("{File_Name}",CurrentFile).replace("{Months_Vacant}", CurrentMonths) 
    MailApp.sendEmail(CurrentEmail,Subject,messagebody)
  }
}
