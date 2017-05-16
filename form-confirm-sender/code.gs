function onFormSubmit() {
  var sheetId   = "";
  var book = SpreadsheetApp.openById(sheetId);
  var sheet = book.getActiveSheet();
  //var sheet = book.getSheetByName(sheetName);
  var memName = getMemberName(sheet);
  var memEmail = getMemberEmail(sheet);
  var isMember = getIsMember(memEmail,sheet);
  var CourseSel = getCourseSel(sheet,isMember);
  var memberChoice = getMemberChoice(sheet);
  var paymentLink = getPaymentLink(memberChoice,isMember);
  var message = HTMLBody(memName,isMember,CourseSel,memberChoice,paymentLink);
  var ts = getTimeStamp(sheet);
  Logger.log(message);
  emailSender(message,ts,memEmail);
}

function emailSender(message,ts,memEmail) {
  //memEmail = "askbmes@gmail.com";
  var subject = "Payment due for BMES TECH Month <"+ts+">";
  Logger.log(message);
  MailApp.sendEmail(memEmail, subject, "", {
    htmlBody: message
    //cc: "",
    //bcc: ""
  });
}

function HTMLBody(memName,IsMember,CourseSel,memberChoice,paymentLink){
  var message="";
  message += "<div id=\"test\" align=\"center\">";
  message += "<img src=\"http:\/\/clubs.ntu.edu.sg\/bmes\/img\/BMES_logo.jpg\" alt=\"\" width=\"252\" height=\"180\"><br><br>";
  message += "<\/div>";
  message += "Dear "+memName+"<br><br>";
  message += IsMember+"<br><br>";
  message += "Thank you for your interest in BMES Tech Month!<br><br>";
  message += "We have noted you have selected the following workshops during your registration:<br>"
  message += "<ul>"+CourseSel+"</ul>";
  message += "Payment link: <a href=\""+paymentLink+"\">"+paymentLink+"</a>.<br><br>";
  message += "Once you have made your payment, you will receive confirmation from us via email.<br><br>";
  message += "Webmaster<br>";
  message += "Biomedical Engineering Society (Student Chapter)<br>";
  message += "Nanyang Technological University";
  return message;
}

function getPaymentLink(memberChoice,isMember){
  var paymentLink="";
  if(isMember == "Your BMES membership has been sucessfully verified."){
    if(memberChoice == "<b>individual courses</b>"){
      paymentLink="";// add payment link
    } else if(memberChoice == "<b>packages</b>"){
      paymentLink="";
    } else{
      paymentLink="";
    }
  } else{
    if(memberChoice == "<b>individual courses</b>"){
      paymentLink="";
    } else if(memberChoice == "<b>packages</b>"){
      paymentLink="";
    } else{
      paymentLink="";
    }
  }
  return paymentLink;
}

//function to send using web API
/*function emailSender(memName,IsMember,CourseSel){
  var url = "";
  var data = { "memName":memName,
              "isMember":IsMember,
              "CourseSel":CourseSel};
  var options = { "method":"POST","payload":data};
  var response = UrlFetchApp.fetch(url ,options);
Logger.log(response);
}*/

function getMemberName(sheet){
  var memName = sheet.getRange(sheet.getLastRow(), 2).getValue();
  return memName;
}

function getCourseSel(sheet,IsMember){
  var CourseSel = "";
  if(sheet.getRange(sheet.getLastRow(), 8).getValue() == "Select individual courses"){
    CourseSel = sheet.getRange(sheet.getLastRow(), 10).getValue();
  } else if(sheet.getRange(sheet.getLastRow(), 8).getValue() == "Select from different packages") {
    CourseSel = sheet.getRange(sheet.getLastRow(), 9).getValue();
  } else{
    CourseSel="All Star Package [Early bird]";
  }
  var courseSelArray = new Array();
  courseSelArray = CourseSel.split(", ");
  var courseStr=""
  for (var i = 0; i < courseSelArray.length; i++) {
    courseStr += "<li>"+courseSelArray[i]+" </li>";
  }
  return courseStr;
}

function getMemberChoice(sheet){
  var memberChoice = "";
  if(sheet.getRange(sheet.getLastRow(), 8).getValue()=="Select from different packages"){
    memberChoice = "<b>packages</b>";
  } else if(sheet.getRange(sheet.getLastRow(), 8).getValue()=="Select individual courses"){
    memberChoice = "<b>individual courses</b>";
  } else{
    memberChoice = "<b>All Star Package</b>";
  }
  return memberChoice;
}

function getMemberEmail(sheet){
  var memEmail = sheet.getRange(sheet.getLastRow(), 4).getValue();
  return memEmail;
}

function getIsMember(memEmail,sheet){
  var mSheetId   = "";
  var mbook = SpreadsheetApp.openById(mSheetId);
  var msheet = mbook.getSheetByName("2016/17");
  var isMember = "<span style=\"color: red\">Your BMES membership cannot be verified. If there has been a mistake, please reply to this email to contact us. If not, please proceed to make payment as a non-member.</span>";
  for (var x = 1; x <= msheet.getLastRow(); x++) {
    if(memEmail.toString().toLowerCase() == msheet.getRange(x, 3).getValue().toString().toLowerCase()){
      isMember = "Your BMES membership has been sucessfully verified.";
    }
  }
  return isMember;
}

function checkEmailQuota(){
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
}

function getTimeStamp(sheet){
  var ts = sheet.getRange(sheet.getLastRow(), 1).getValue();
  return ts;
}

function autoFormDisable(){
  var form = FormApp.openById("");
  var defaultClosedFor = form.getCustomClosedFormMessage();
  var date = new Date();
  var deadline = new Date(2017,4,29,0,0);
  if (date>=deadline){
    form.setCustomClosedFormMessage("The early bird offer has ended. Please register using the new link on our website. Visit http://clubs.ntu.edu.sg/bmes/events.html");
    form.setAcceptingResponses(false);
  } else{
    form.setAcceptingResponses(true);
  }
}