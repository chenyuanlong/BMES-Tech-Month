function doGet(e) {
  return HtmlService.createTemplateFromFile('receipt').evaluate();
}

function processForm(myForm) {
  var alreadySent = false;
  var sheetId   = "";//google sheet id
  var book = SpreadsheetApp.openById(sheetId);
  var sheet = book.getActiveSheet();
  var rowNum = myForm.select;
  var memName = sheet.getRange(rowNum, 2).getValue();
  var memEmail = sheet.getRange(rowNum, 4).getValue();
  var isMember = getIsMember(memEmail,sheet);
  var CourseSel = getCourseSel(rowNum,sheet,isMember);
  var message = HTMLBody(memName,CourseSel);
  var HPNo = sheet.getRange(rowNum, 6).getValue();
  var totalPrice = getTotalPrice(rowNum,sheet,isMember);
  if(myForm.select!="error" && eReceiptSent(HPNo)){
    Logger.log("Message sent already.");
    alreadySent = true;
  }
  else if(myForm.select!="error" && !eReceiptSent(HPNo)){
    Logger.log("Message not sent before.");
    Logger.log(memEmail+"\n"+message);
    emailSender(message,memEmail,memName);
    addtoChecklist(rowNum,totalPrice);
    alreadySent = false;
  }
  return alreadySent;
}

function returnMemData(myForm){
  if(myForm!="error"){
    var sheetId   = "";
    var book = SpreadsheetApp.openById(sheetId);
    var sheet = book.getActiveSheet();
    var rowNum = myForm;
    var memEmail = sheet.getRange(rowNum, 4).getValue();
    var isMember = getIsMember(memEmail,sheet);
    var CourseSel = getCourseSelWebApp(rowNum,sheet,isMember);
    return CourseSel;
  }
}

function emailSender(message,memEmail,memName) {
  //memEmail = "askbmes@gmail.com";
  var subject = "Payment Receipt for BMES Tech Month";
  MailApp.sendEmail(memEmail, subject, "", {
    htmlBody: message
    //cc: "", //CC & BCC if needed
    //bcc: ""
  });
}

function HTMLBody(memName,CourseSel){
  var message="";
  message += "<div id=\"test\" align=\"center\">";
  message += "<img src=\"http:\/\/clubs.ntu.edu.sg\/bmes\/img\/BMES_logo.jpg\" alt=\"\" width=\"252\" height=\"180\"><br><br>";
  message += "<\/div>";
  message += "Dear "+memName+"<br><br>";
  message += "Thank you for registering with us. We have received your payment for the following workshops.<br><br>";
  message += "<table style=\"background-color:#F7FDFA;border-collapse:collapse;color:#000\">";
  message += "  <tr>";
  message += "    <th style=\"background-color:#26ADE4;color:white;width:33%;padding:10px 20px;border:0;text-align:center\">S\/N<\/th>";
  message += "    <th style=\"background-color:#26ADE4;color:white;width:33%;padding:10px 20px;border:0;text-align:center\">Workshop<\/th>";
  message += "    <th style=\"background-color:#26ADE4;color:white;width:33%;padding:10px 20px;border:0;text-align:center\">Price<\/th>";
  message += "  <\/tr>";
  message += CourseSel;
  /*message += "  <tr>";
  message += "    <td style=\"background-color:#26ADE4;color:white;width:33%;padding:10px 20px;border:0;text-align:center\">Total<\/td>";
  message += "    <td style=\"background-color:#26ADE4;color:white;width:33%;padding:10px 20px;border:0;text-align:center\"><\/td>";
  message += "    <td style=\"background-color:#26ADE4;color:white;width:33%;padding:10px 20px;border:0;text-align:center\">$3,000,000<\/td>";
  message += "  <\/tr>";*/
  message += "<\/table><br>";
  message += "Please note that you are only allowed to attend the workshop that you have paid for.<br>";
  message += "If you have any enquries, please do not hesitate to contact us.<br><br>";
  message += "Cheers<br><br>";
  message += "<b>Financial Controller</b><br>";
  message += "Biomedical Engineering Society (Student Chapter)<br>";
  message += "Nanyang Technological University";
  return message;
}

function getCourseSel(rowNum,sheet,isMember){
  var CourseSel = "";
  var packageType = "";
  var price = "";
  var totalPrice = 0;
  if(sheet.getRange(rowNum, 8).getValue() == "Select individual courses"){
    packageType = sheet.getRange(rowNum, 8).getValue();
    CourseSel = sheet.getRange(rowNum, 10).getValue();
  } else if(sheet.getRange(rowNum, 8).getValue() == "Select from different packages") {
    packageType = sheet.getRange(rowNum, 8).getValue();
    CourseSel = sheet.getRange(rowNum, 9).getValue();
  } else{
    packageType = "All Star Package [Early bird]";
    CourseSel="All Star Package [Early bird]";
  }
  var courseSelArray = new Array();
  courseSelArray = CourseSel.split(", ");
  var courseStr="";
  for (var i = 0; i < courseSelArray.length; i++) {
   if(isMember && packageType == "Select individual courses"){
     switch (courseSelArray[i]) {
       case 'Arduino Basic':
         price = "5.25";
         break;
       case 'Arduino Advanced':
         price = "5.25";
         break;
       case 'CREO Basic':
         price = "3.75";
         break;
       case 'CREO Advanced':
         price = "3.75";
         break;
       case 'VBA':
         price = "3.75";
         break;
       case 'Web Development':
         price = "3.75";
         break;
       default:
         price = "ERROR."
     }
   } else if(isMember && packageType == "Select from different packages"){
      switch (courseSelArray[i]) {
        case 'Arduino Complete':
          price = "9";
          break;
        case 'CREO Complete':
          price = "6";
          break;
        case 'Skills Future':
          price = "6";
          break;
        default:
          price = "ERROR."
      }
   } else if(isMember && packageType == "All Star Package [Early bird]"){
       price = "18";
    } else if(!isMember && packageType == "Select individual courses"){
      switch (courseSelArray[i]) {
       case 'Arduino Basic':
         price = "7";
         break;
       case 'Arduino Advanced':
         price = "7";
         break;
       case 'CREO Basic':
         price = "5";
         break;
       case 'CREO Advanced':
         price = "5";
         break;
       case 'VBA':
         price = "5";
         break;
       case 'Web Development':
         price = "5";
         break;
       default:
         price = "ERROR."
      }
    } else if(!isMember && packageType == "Select from different packages"){
      switch (courseSelArray[i]) {
        case 'Arduino Complete':
          price = "12";
          break;
        case 'CREO Complete':
          price = "8";
          break;
        case 'Skills Future':
          price = "8";
          break;
        default:
          price = "ERROR."
      }
    } else{
      price = "24";
    }
    totalPrice += Number(price); 
    courseStr += "  <tr>";
    courseStr += "    <td style=\"padding:10px 20px;border:0;text-align:center;border-bottom:1px dotted #BDB76B;\">"+Number(i+1)+"<\/td>";
    courseStr += "    <td style=\"padding:10px 20px;border:0;text-align:center;border-bottom:1px dotted #BDB76B;\">"+courseSelArray[i]+"<\/td>";
    courseStr += "    <td style=\"padding:10px 20px;border:0;text-align:center;border-bottom:1px dotted #BDB76B;\">$"+price+"<\/td>";
    courseStr += "  <\/tr>";
  }// end for loop
  courseStr += "  <tr>";
  courseStr += "    <td style=\"background-color:#26ADE4;color:white;width:33%;padding:10px 20px;border:0;text-align:center\">Total<\/td>";
  courseStr += "    <td style=\"background-color:#26ADE4;color:white;width:33%;padding:10px 20px;border:0;text-align:center\"><\/td>";
  courseStr += "    <td style=\"background-color:#26ADE4;color:white;width:33%;padding:10px 20px;border:0;text-align:center\">$"+totalPrice+"<\/td>";
  courseStr += "  <\/tr>";
  return courseStr;
}

function getCourseSelWebApp(rowNum,sheet,isMember){
  var CourseSel = "";
  if(sheet.getRange(rowNum, 8).getValue() == "Select individual courses"){
    CourseSel = sheet.getRange(rowNum, 10).getValue();
  } else if(sheet.getRange(rowNum, 8).getValue() == "Select from different packages") {
    CourseSel = sheet.getRange(rowNum, 9).getValue();
  } else{
    CourseSel="All Star Package [Early bird]";
  }
  var courseSelArray = new Array();
  courseSelArray = CourseSel.split(", ");
  var courseStr="";
  for (var i = 0; i < courseSelArray.length; i++) {
    //courseStr += courseSelArray[i]+"\n";
    courseStr += courseSelArray[i] + "<br>";
  }
  courseStr += "<br>Phone No: "+sheet.getRange(rowNum, 6).getValue() + "<br>";
  courseStr += "<br>Email: "+sheet.getRange(rowNum, 4).getValue() + "<br><br>";
  var totalPrice = getTotalPrice(rowNum,sheet,isMember);
  courseStr += "Total Amount Due: $"+totalPrice+ "<br><br>";
  courseStr += " Remaining email quota: "+checkEmailQuota()+ "<br><br>";
  return courseStr;
}

function getIsMember(memEmail,sheet){  
  var mSheetId   = "";
  var mbook = SpreadsheetApp.openById(mSheetId);
  var msheet = mbook.getSheetByName("2016/17");
  var isMember = false;
  for (var x = 1; x <= msheet.getLastRow(); x++) {
    if(memEmail.toString().toLowerCase() == msheet.getRange(x, 3).getValue().toString().toLowerCase()){
      isMember = true;
    }
  }
  return isMember;
}

function getTotalPrice(rowNum,sheet,isMember){
  var totalPrice = 0;
  if(sheet.getRange(rowNum, 8).getValue() == "Select individual courses"){
    packageType = sheet.getRange(rowNum, 8).getValue();
    CourseSel = sheet.getRange(rowNum, 10).getValue();
  } else if(sheet.getRange(rowNum, 8).getValue() == "Select from different packages") {
    packageType = sheet.getRange(rowNum, 8).getValue();
    CourseSel = sheet.getRange(rowNum, 9).getValue();
  } else{
    packageType = "All Star Package [Early bird]";
    CourseSel="All Star Package [Early bird]";
  }
  var courseSelArray = new Array();
  courseSelArray = CourseSel.split(", ");
  var courseStr="";
  for (var i = 0; i < courseSelArray.length; i++) {
   if(isMember && packageType == "Select individual courses"){
     switch (courseSelArray[i]) {
       case 'Arduino Basic':
         price = "5.25";
         break;
       case 'Arduino Advanced':
         price = "5.25";
         break;
       case 'CREO Basic':
         price = "3.75";
         break;
       case 'CREO Advanced':
         price = "3.75";
         break;
       case 'VBA':
         price = "3.75";
         break;
       case 'Web Development':
         price = "3.75";
         break;
       default:
         price = "ERROR."
     }
   } else if(isMember && packageType == "Select from different packages"){
      switch (courseSelArray[i]) {
        case 'Arduino Complete':
          price = "9";
          break;
        case 'CREO Complete':
          price = "6";
          break;
        case 'Skills Future':
          price = "6";
          break;
        default:
          price = "ERROR."
      }
   } else if(isMember && packageType == "All Star Package [Early bird]"){
       price = "18"
    } else if(!isMember && packageType == "Select individual courses"){
      switch (courseSelArray[i]) {
       case 'Arduino Basic':
         price = "7";
         break;
       case 'Arduino Advanced':
         price = "7";
         break;
       case 'CREO Basic':
         price = "5";
         break;
       case 'CREO Advanced':
         price = "5";
         break;
       case 'VBA':
         price = "5";
         break;
       case 'Web Development':
         price = "5";
         break;
       default:
         price = "ERROR."
      }
    } else if(!isMember && packageType == "Select from different packages"){
      switch (courseSelArray[i]) {
        case 'Arduino Complete':
          price = "12";
          break;
        case 'CREO Complete':
          price = "8";
          break;
        case 'Skills Future':
          price = "8";
          break;
        default:
          price = "ERROR."
      }
    } else{
      price = "24";
    }
    totalPrice += Number(price);
  }
  return  totalPrice;
}

function checkEmailQuota(){
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
  return emailQuotaRemaining;
}

function addtoChecklist(rowNum,totalPrice){
  var msheetId   = "";
  var mbook = SpreadsheetApp.openById(msheetId);
  var msheet = mbook.getActiveSheet();
  var sheetId   = "";
  var book = SpreadsheetApp.openById(sheetId);
  var sheet = book.getActiveSheet();
  var lastrow = sheet.getLastRow()+1;
  sheet.getRange(lastrow,1).setValue(msheet.getRange(rowNum, 2).getValue());
  sheet.getRange(lastrow,2).setValue(msheet.getRange(rowNum, 6).getValue());
  sheet.getRange(lastrow,3).setValue(totalPrice);
}

function eReceiptSent(HPNo){
  var receiptSent =false;
  var sheetId   = "";
  var book = SpreadsheetApp.openById(sheetId);
  var sheet = book.getActiveSheet();
  for(var i = 1; i <= sheet.getLastRow(); i++){
    if(sheet.getRange(i,2).getValue()==HPNo){
      receiptSent = true;
    }
  }
  return receiptSent;
}