var SHEET_NAME = "Personnes intéressées";
var TEMPLATE_SHEET_NAME= "Email Templates";
var CONFIG_SHEET_NAME="Config";
var TIME_RANGE_MINUTES=15;
var SPREADSHEET_ID="1-wZOfNuKbSXtpw1zYJW9BBonhm8vDrcgomfJA6axVPk";
var MAX_ROW_TO_SCAN_FROM_END=50;

function doPost(e) {
  var spreadsheet=SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  var dataJson=JSON.parse(e["parameters"]["data.json"]);
  var date= getDateYyyyMmDd();
  var time= getTimeHhMmSs();
  var civilite=dataJson["civilite"];
  var lastName = dataJson["nom_de_famille"];
  var email = dataJson["email"];
  var variant = e.parameters.variant;
  var page_uuid = e.parameter.page_id;
  var gclid = dataJson["gclid"];
  var javascript_prepopulated_value = dataJson["javascript_prepopulated_value"];
  var stage = 1;
  var lastRowNum=sheet.getLastRow()+1;
  
  sheet.getRange("A"+lastRowNum).setValue(date);
  sheet.getRange("B"+lastRowNum).setValue(time);
  sheet.getRange("C"+lastRowNum).setValue(civilite);
  sheet.getRange("E"+lastRowNum).setValue(lastName);
  sheet.getRange("F"+lastRowNum).setValue(email);
  sheet.getRange("O"+lastRowNum).setValue(variant);
  sheet.getRange("P"+lastRowNum).setValue(page_uuid);
  sheet.getRange("Q"+lastRowNum).setValue(gclid);
  sheet.getRange("R"+lastRowNum).setValue(javascript_prepopulated_value);
  sheet.getRange("U"+lastRowNum).setValue(stage);
  sheet.getRange("V"+lastRowNum).setValue(new Date().getTime());
  
}
function getDateYyyyMmDd() {
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; //January is 0!
  
  var yyyy = today.getFullYear();
  if(dd<10){
    dd='0'+dd
  } 
  if(mm<10){
    mm='0'+mm
  } 
  var today = yyyy+'/'+mm+'/'+dd;
  Logger.log("today: "+today);
  return today;
}
function getTimeHhMmSs() {
  var now = new Date();
  var hh = now.getHours();
  var mm = now.getMinutes();
  var ss=now.getSeconds();
  
  var timenow = hh+':'+mm+':'+ss;
  Logger.log("today: "+timenow);
  return timenow;
}
function checkAndSendEmail1() {
  var spreadsheet=SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  var configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  var templateSheet = spreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);
  var spreadsheet_link_client = configSheet.getRange("B6").getValue();
  var nom_client = configSheet.getRange("B2").getValue();
  var recipient1=configSheet.getRange("B3").getValue();
  var recipient2=configSheet.getRange("B4").getValue();
  var recipient3=configSheet.getRange("B5").getValue();
  
  var fromName=configSheet.getRange("B9").getValue();
  
  var startRowNum = sheet.getLastRow();
  var rowsScanned=0;
  // loop thru each row
  for(var i = startRowNum; i >= 2; i--) {
    var status = sheet.getRange("U"+i).getValue();
    var rowNum=i;
    
    if(status == 1 || status == 1.0) {
      var timestamp = parseInt(String(sheet.getRange("V"+i).getValue()), 10);
      var civilite = sheet.getRange("C"+rowNum).getValue();
      var nom_de_famille = sheet.getRange("E"+rowNum).getValue();
      var email = sheet.getRange("F"+rowNum).getValue();
      
      
      var mailSub = String(templateSheet.getRange("C3").getValue()).replace("{civilite}",civilite).replace("{nom_de_famille}", nom_de_famille)
      .replace("{ConfigSheet/nom_client}", nom_client).replace("{email}", email).replace("{ConfigSheet/spreadsheet_link_client}", spreadsheet_link_client);
      
      var mailBody = String(templateSheet.getRange("D3").getValue()).replace("{civilite}",civilite).replace("{nom_de_famille}", nom_de_famille)
      .replace("{ConfigSheet/nom_client}", nom_client).replace("{email}", email).replace("{ConfigSheet/spreadsheet_link_client}", spreadsheet_link_client);
      
      GmailApp.sendEmail(recipient1+","+recipient2+","+recipient3, mailSub, mailBody, {name: fromName});
      sheet.getRange("U"+i).setValue(2);
      sheet.getRange("J"+i).setValue("Tarifs");
    }
    rowsScanned++;
    if(rowsScanned==MAX_ROW_TO_SCAN_FROM_END) {
      break;
    }
  }
}

function checkTimeInRange(formerTime) {
  var curTime=new Date().getTime();
  var interval = TIME_RANGE_MINUTES*60*1000; // ms
  var difference = curTime-formerTime;
  // check if time is not in future
  if(difference >= 0) { 
    if(difference >= interval) {
      return true;
    }
  }
  return false;
}
