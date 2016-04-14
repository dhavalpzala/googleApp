var SHEET_NAME = "Personnes intéressées";
var TEMPLATE_SHEET_NAME= "Email Templates";
var CONFIG_SHEET_NAME="Config";
var MAX_ROW_TO_SCAN_FROM_END=50;

function getHeadersJSON(sheetObj) {  
  var columnsCount = sheetObj.getLastColumn(), 
      headersRange = sheetObj.getRange(1,1,1,columnsCount).getValues(),
      headersJSON = {};
  
  for (var headerIndex = 0; headerIndex < headersRange[0].length; headerIndex++) {
    headersJSON[headersRange[0][headerIndex]] = headerIndex + 1;
  }
  
  return headersJSON;
};

function doPost(e) {
  
  var spreadsheet=SpreadsheetApp.openById("1-wZOfNuKbSXtpw1zYJW9BBonhm8vDrcgomfJA6axVPk");
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  var headersJSON = getHeadersJSON(sheet);
  
  var dataJson=JSON.parse(e["parameters"]["data.json"]);
  var javascript_prepopulated_value = dataJson["javascript_prepopulated_value"];
  var message = dataJson["d\u00e9crivez_votre_souhait"];
  var telephone = dataJson["phone"];
  var rowNum = Math.round(foundJPVinSheet(sheet, String(javascript_prepopulated_value), headersJSON['javascript_prepopulated_value(from unbounce=timestamp)']));
  if(rowNum==0) {
    rowNum=sheet.getLastRow()+1;
  }
  sheet.getRange(rowNum, headersJSON['Message']).setValue(message);
  sheet.getRange(rowNum, headersJSON['Téléphone']).setValue(telephone);
  sheet.getRange(rowNum, headersJSON['Action']).setValue("A rappeler");
  sheet.getRange(rowNum, headersJSON['status']).setValue(2);
  
  
  // code to now send email
  var configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  var templateSheet = spreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);
  var stage = sheet.getRange(rowNum, headersJSON['status']).getValue();
  
  var spreadsheet_link_client = configSheet.getRange("B6").getValue();
  var civilite = sheet.getRange(rowNum, headersJSON['Civilité']).getValue();
  var nom_de_famille = sheet.getRange(rowNum, headersJSON['Nom de Famille(Last Name)']).getValue();
  var nom_client = configSheet.getRange("B2").getValue();
  var email = sheet.getRange(rowNum, headersJSON['Email']).getValue();
  var recipient1=configSheet.getRange("B3").getValue();
  var recipient2=configSheet.getRange("B4").getValue();
  var recipient3=configSheet.getRange("B5").getValue();
  var cc=configSheet.getRange("B10").getValue();
  var bcc=configSheet.getRange("B11").getValue();
  
  var fromName = String(configSheet.getRange("B9").getValue());
  
  var phone = sheet.getRange(rowNum, headersJSON['Téléphone']).getValue();
  var message = sheet.getRange(rowNum, headersJSON['Message']).getValue();
  
  
  var mailSub = String(templateSheet.getRange("C2").getValue()).replace("{civilite}",civilite).replace("{nom_de_famille}", nom_de_famille).replace("{phone}",phone)
  .replace("{ConfigSheet/nom_client}", nom_client).replace("{email}", email).replace("{souhait}",message).replace("{ConfigSheet/spreadsheet_link_client}", spreadsheet_link_client);
  
  var mailBody = String(templateSheet.getRange("D2").getValue()).replace("{civilite}",civilite).replace("{nom_de_famille}", nom_de_famille).replace("{phone}",phone)
  .replace("{ConfigSheet/nom_client}", nom_client).replace("{email}", email).replace("{souhait}",message).replace("{ConfigSheet/spreadsheet_link_client}", spreadsheet_link_client);
  
  GmailApp.sendEmail(recipient1+","+recipient2+","+recipient3, mailSub, mailBody, {name: fromName, cc: cc, bcc: bcc});
}

function foundJPVinSheet(sheet, jPVToFind, jpvColumnIndex) {
  var lastRowNum=sheet.getLastRow();
  var rowScanned = 0;
  for(var i=lastRowNum;i>=2;i--) {
    var jpv=sheet.getRange(i, jpvColumnIndex).getValue();
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
