function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gmaven Mail Merge')
      .addItem('Send Emails', 'sendEmail')
      .addToUi();
}
 





function sendEmail() {
var Property_name=2;
var Address=3;
var Subrurb=4;
var City=5;	
var fileName=7;
var	Province=6;
var unit_type=8;	
var unit_sub_type=9;
var	building_name=10;
var	floor_no=11;
var	unit_no=12;
var	unit_gla=13;
var	available_type=14;
var	available_date=15;
var	rent_tba=16;
var	gross_rent=17;		
var Months_vacant= 20;	
var Business_Name=21;
var Contact_name=22;
var	Contact_surname=23;	
var Contact_email=24;
var Email_template=25;	
var send=33;
var email_Sent=34;
var sent_Date=35;


var emailtemp=HtmlService.createTemplateFromFile("Email.html");

var worksheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");

var worksheet2= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Template");

var data= worksheet.getRange('A2:AF'+ worksheet.getLastRow()).getValues();




var Template=worksheet2.getRange('B4').getValue();


 var SenderName= worksheet2.getRange("B2").getValue()

var increment=2;


data.forEach(function(row){

  // var EnhanceTemplate=Template.replace("{Name}",row[Contact_name]).replace("{File_Name}",row[fileName]).replace("{Months_Vacant}",row[Months_vacant]).replace(/\n/g,'<p>/<p>' )

  var EMAIL_SENT_Text = 'Yes';

  var check = worksheet.getRange(increment,send).getValue()


  var date = Utilities.formatDate(new Date(), "GMT+2", "dd/MM/yyyy-HH:mm")

  if(check==true){
  emailtemp.fn=row[Contact_name];
  emailtemp.filename=row[fileName];
  emailtemp.MonthVacant=row[Months_vacant];
  var sub =worksheet2.getRange('B3').getValue().replace("<<business name>>",row[Business_Name]);

  var finalmessage= emailtemp.evaluate().getContent();


  GmailApp.sendEmail(row[Contact_email],sub,
  "Your Browser Doesnt Supports HTML",
  {name:SenderName,htmlBody:finalmessage})


    // GmailApp.sendEmail(row[Contact_email],sub,finalmessage)
  
    worksheet.getRange(increment,email_Sent).setValue(EMAIL_SENT_Text).setFontColor("red")
    worksheet.getRange(increment,sent_Date).setValue(date).setFontColor("red")
  }


 increment++;


})

}
