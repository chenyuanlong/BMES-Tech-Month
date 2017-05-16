function doGet(e) {
  return HtmlService.createTemplateFromFile('handout').evaluate();
}

function processForm(myForm) {
  var alreadySent = false;
  var sheetId   = ""; //google sheet id
  var book = SpreadsheetApp.openById(sheetId);
  var sheet = book.getActiveSheet();
  var rowNum = myForm.select;
  var courseSel = myForm.select2;
  var memName = sheet.getRange(rowNum, 2).getValue();
  var memEmail = sheet.getRange(rowNum, 4).getValue();
  var message = HTMLBody(memName,courseSel);
  var HPNo = sheet.getRange(rowNum, 6).getValue();
  Logger.log(memEmail+"\n"+message);
  if(myForm.select!="error" && nameInList(HPNo)){
    if(courseSel=="creoB"){
      alreadySent = arduinoBSent(HPNo);
      if(!alreadySent){
        emailSender(message,memEmail,memName,courseSel);
        Logger.log("CREO basic sent");
      }
    } else{
      alreadySent = arduinoASent(HPNo);
      if(!alreadySent){
        emailSender(message,memEmail,memName);
        Logger.log("CREO advanced sent");
      }
    }
  }
  else if(myForm.select!="error" && !nameInList(HPNo)){
    alreadySent = false;
    addtoChecklist(rowNum,courseSel);
    emailSender(message,memEmail,memName,courseSel);
  }
  return alreadySent;
}

function returnMemData(myForm){
  if(myForm!="error"){
    var sheetId   = ""; //google sheet id for member list
    var book = SpreadsheetApp.openById(sheetId);
    var sheet = book.getActiveSheet();
    var rowNum = myForm;
    var memEmail = sheet.getRange(rowNum, 4).getValue();
    var CourseSel = getCourseSelWebApp(rowNum,sheet);
    return CourseSel;
  }
}

function emailSender(message,memEmail,memName,courseSel) {
  //memEmail = ""; //for sending to self for testing
  var subject = "";
  if(courseSel=="creoB"){
    subject = "BMES Tech Month - CREO Basic Lesson Details";
  }else{
    subject = "BMES Tech Month - CREO Advanced Lesson Details";
  }
  MailApp.sendEmail(memEmail, subject, "", {
    htmlBody: message
    //cc: "", //CC & BCC if needed
    //bcc: ""
  });
}

function HTMLBody(memName,courseSel){
  var message="";
  if(courseSel=="creoB"){
    message += "<div id=\"test\" align=\"center\">";
    message += "<img src=\"http:\/\/clubs.ntu.edu.sg\/bmes\/techmonth\/VenueCreoBasic.jpg\" width=\"624\" height=\"327\"><br><br>";
    message += "<\/div>";
    message += "Dear "+memName+"<br><br>";
    message += "Thank you for registering for the CREO Basic Class. Please take note of the following details:<br><br>";
    message += "<b>Workshop</b>: CREO Basic<br>";
    message += "<b>Venue</b>: SCBE, N1.2-B4-02 Computer Lab<br>";
    message += "<b>Map</b>: <a href=\"http://maps.ntu.edu.sg/maps#q:N1.2-B4-02\">http://maps.ntu.edu.sg/maps#q:N1.2-B4-02</a><br>";
    message += "<b>Date</b>: 27 May 2017, Saturday<br>";
    message += "<b>Time</b>: 9:30am - 12:30pm<br>";
    message += "<span style=\"color:red\"><i>Registration starts from 9:00am, please be present before the lesson starts.</i></span><br>";
    message += "<span style=\"color:red\"><i>Important: Please remember to bring along your laptop for this workshop.</i></span><br><br>";
    message += "Please download a copy of your lesson handout from the link below and feel free to print or bring along a soft copy according to your preference.<br>";
    message += "<a href=\"http://clubs.ntu.edu.sg/bmes/techmonth/\">http://clubs.ntu.edu.sg/bmes/techmonth/</a><br><br>";
    message += "Cheers<br>";
    message += "<span style=\"color:blue\"><i><b>askBMES</b></i></span><br>";
  } else{
    message += "<div id=\"test\" align=\"center\">";
    message += "<img src=\"http:\/\/clubs.ntu.edu.sg\/bmes\/techmonth\/VenueCreoAdvanced.jpg\" width=\"624\" height=\"327\"><br><br>";
    message += "<\/div>";
    message += "Dear "+memName+"<br><br>";
    message += "Thank you for registering for the CREO Advanced Class. Please take note of the following details:<br><br>";
    message += "<b>Workshop</b>: CREO Advanced<br>";
    message += "<b>Venue</b>: SCBE, N1.2-B4-02, Computer Lab<br>";
    message += "<b>Map</b>: <a href=\"http://maps.ntu.edu.sg/maps#q:N1.2-B4-02\">http://maps.ntu.edu.sg/maps#q:N1.2-B4-02</a><br>";
    message += "<b>Date</b>: 27 May 2017, Saturday<br>";
    message += "<b>Time</b>: 2:00pm - 5:00pm<br>";
    message += "<span style=\"color:red\"><i>Registration starts from 1:30pm, please be present before the lesson starts.</i></span><br>";
    message += "<span style=\"color:red\"><i>Important: Please remember to bring along your laptop for this workshop.</i></span><br><br>";
    message += "Please download a copy of your lesson handout from the link below and feel free to print or bring along a soft copy according to your preference.<br>";
    message += "<a href=\"http://clubs.ntu.edu.sg/bmes/techmonth/\">http://clubs.ntu.edu.sg/bmes/techmonth/</a><br><br>";
    message += "Cheers<br>";
    message += "<span style=\"color:blue\"><i><b>askBMES</b></i></span><br>";
  }
  return message;
}



function getCourseSelWebApp(rowNum,sheet){
  var courseStr = "";
  courseStr += "<br>Phone No: "+sheet.getRange(rowNum, 6).getValue() + "<br>";
  courseStr += "<br>Email: "+sheet.getRange(rowNum, 4).getValue() + "<br><br>";
  courseStr += " Remaining email quota: "+checkEmailQuota()+ "<br><br>";
  return courseStr;
}

function checkEmailQuota(){
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
  return emailQuotaRemaining;
}

function addtoChecklist(rowNum,courseSel){
  var msheetId   = ""; //google sheet id for member list 
  var mbook = SpreadsheetApp.openById(msheetId);
  var msheet = mbook.getActiveSheet();
  var sheetId   = "";//google sheet id
  var book = SpreadsheetApp.openById(sheetId);
  var sheet = book.getActiveSheet();
  var lastrow = sheet.getLastRow()+1;
  sheet.getRange(lastrow,1).setValue(msheet.getRange(rowNum, 2).getValue());
  sheet.getRange(lastrow,2).setValue(msheet.getRange(rowNum, 6).getValue());
  if(courseSel=="arduinoB"){
    sheet.getRange(lastrow,7).setValue("CREO Basic");
  } else{
    sheet.getRange(lastrow,8).setValue("CREO Advanced");
  }
}

function nameInList(HPNo){
  var nameInList =false;
  var sheetId   = "";//google sheet id
  var book = SpreadsheetApp.openById(sheetId);
  var sheet = book.getActiveSheet();
  for(var i = 1; i <= sheet.getLastRow(); i++){
    if(sheet.getRange(i,2).getValue()==HPNo){
      nameInList = true;
    }
  }
  return nameInList;
}

function arduinoBSent(HPNo){
  var rowNum = 0;
  var vbaSent =false;
  var sheetId   = "";//google sheet id
  var book = SpreadsheetApp.openById(sheetId);
  var sheet = book.getActiveSheet();
  for(var i = 1; i <= sheet.getLastRow(); i++){
    if(sheet.getRange(i,2).getValue()==HPNo){
      rowNum = i;
    }
  }
  if(sheet.getRange(rowNum,7).getValue()=="CREO Basic"){
    vbaSent = true;
  } else{
    sheet.getRange(rowNum,7).setValue("CREO Basic");
  }
  return vbaSent;
}

function arduinoASent(HPNo){
  var rowNum = 0;
  var htmlSent =false;
  var sheetId   = "";//google sheet id
  var book = SpreadsheetApp.openById(sheetId);
  var sheet = book.getActiveSheet();
  for(var i = 1; i <= sheet.getLastRow(); i++){
    if(sheet.getRange(i,2).getValue()==HPNo){
      rowNum = i;
    }
  }
  if(sheet.getRange(rowNum,8).getValue()=="CREO Advanced"){
    htmlSent = true;
  } else{
    sheet.getRange(rowNum,8).setValue("CREO Advanced");
  }
  return htmlSent;
}