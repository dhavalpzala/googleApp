var SHEET_NAME = "Personnes intéressées";
var TEMPLATE_SHEET_NAME= "Email Templates";
var CONFIG_SHEET_NAME="Config";
var MAX_ROW_TO_SCAN_FROM_END=50;
function doPost(e) {
  
  var spreadsheet=SpreadsheetApp.openById("1-wZOfNuKbSXtpw1zYJW9BBonhm8vDrcgomfJA6axVPk");
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  var dataJson=JSON.parse(e["parameters"]["data.json"]);
  var javascript_prepopulated_value = dataJson["javascript_prepopulated_value"];
  var message = dataJson["d\u00e9crivez_votre_souhait"];
  var telephone = dataJson["phone"];
  var rowNum = Math.round(foundJPVinSheet(sheet, String(javascript_prepopulated_value)));
  if(rowNum==0) {
    rowNum=sheet.getLastRow()+1;
  }
  sheet.getRange(rowNum, 9).setValue(message);
  sheet.getRange("G"+rowNum).setValue(telephone);
  sheet.getRange("J"+rowNum).setValue("A rappeler");
  sheet.getRange("U"+rowNum).setValue(2);
  
  
  // code to now send email
  var configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  var templateSheet = spreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);
  var stage = sheet.getRange("U"+rowNum).getValue();
  
  var spreadsheet_link_client = configSheet.getRange("B6").getValue();
  var civilite = sheet.getRange("C"+rowNum).getValue();
  var nom_de_famille = sheet.getRange("E"+rowNum).getValue();
  var nom_client = configSheet.getRange("B2").getValue();
  var email = sheet.getRange("F"+rowNum).getValue();
  var recipient1=configSheet.getRange("B3").getValue();
  var recipient2=configSheet.getRange("B4").getValue();
  var recipient3=configSheet.getRange("B5").getValue();
  
  var fromName = String(configSheet.getRange("B9").getValue());
  
  var phone = sheet.getRange("G"+rowNum).getValue();
  var message = sheet.getRange("I"+rowNum).getValue();
  
  
  var mailSub = String(templateSheet.getRange("C2").getValue()).replace("{civilite}",civilite).replace("{nom_de_famille}", nom_de_famille).replace("{phone}",phone)
  .replace("{ConfigSheet/nom_client}", nom_client).replace("{email}", email).replace("{souhait}",message).replace("{ConfigSheet/spreadsheet_link_client}", spreadsheet_link_client);
  
  var mailBody = String(templateSheet.getRange("D2").getValue()).replace("{civilite}",civilite).replace("{nom_de_famille}", nom_de_famille).replace("{phone}",phone)
  .replace("{ConfigSheet/nom_client}", nom_client).replace("{email}", email).replace("{souhait}",message).replace("{ConfigSheet/spreadsheet_link_client}", spreadsheet_link_client);
  
  GmailApp.sendEmail(recipient1+","+recipient2+","+recipient3, mailSub, mailBody, {name: fromName});
}

function foundJPVinSheet(sheet, jPVToFind) {
  var lastRowNum=sheet.getLastRow();
  var rowScanned = 0;
  for(var i=lastRowNum;i>=2;i--) {
    var jpv=sheet.getRange("R"+i).getValue();
    if(String(jpv).trim()==jPVToFind) { 
      return i;
    }
    rowScanned++;
    if(rowScanned>=MAX_ROW_TO_SCAN_FROM_END) {
      return 0;
    }
  }
  return 0;
}
