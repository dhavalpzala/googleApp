var App = {
  SPREADSHEET_ID: '1Iy6ipllGfW71k__tCZpbG1U1cZc9oYgx7cDlQDXCjRw',
  SHEETS: {
    INTERESTED_PERSONS: 'Personnes intéressées',
    EMAIL_TEMPLATES: 'Email Templates',
    CONFIG: 'Config'
  }
};

function convertRangeToJSONArray(data) {
  var headers = [], 
      JSONArrayObj = [];
  
  for (var headerIndex = 0; headerIndex < data[0].length; headerIndex++) {
    headers.push(data[0][headerIndex]);
  }
  
  for (var rowIndex = 1; rowIndex < data.length; rowIndex++) {
    var obj = {};
    for (var colIndex = 0; colIndex < headers.length; colIndex++) {
      obj[headers[colIndex]] = data[rowIndex][colIndex];
    }
    JSONArrayObj.push(obj);
  }
  
  return JSONArrayObj;
};

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
  var spreadsheet=SpreadsheetApp.openById(App.SPREADSHEET_ID),
      sheet = spreadsheet.getSheetByName(App.SHEETS.INTERESTED_PERSONS),
      headersJSON = getHeadersJSON(sheet),
      JSONData = JSON.parse(e["parameters"]["data.json"]);
  
  if(JSONData) {
    var date= getDateYyyyMmDd(),
        time= getTimeHhMmSs(),
        civilite=JSONData["civilite"],
        lastName = JSONData["nom_de_famille"],
        email = JSONData["email"],
        phoneNo = JSONData["phone"],
        variant = e.parameters.variant,
        page_uuid = e.parameter.page_id,
        gclid = JSONData["gclid"],
        javascript_prepopulated_value = JSONData["javascript_prepopulated_value"],
        stage = 1,
        timeStamp = new Date().getTime(),
        lastRowNum=sheet.getLastRow()+1;
    
    //Create new row with respective values
    sheet.getRange(lastRowNum, headersJSON['Date']).setValue(date);
    sheet.getRange(lastRowNum, headersJSON['Heure']).setValue(time);
    sheet.getRange(lastRowNum, headersJSON['Civilité']).setValue(civilite);
    sheet.getRange(lastRowNum, headersJSON['Nom de Famille(Last Name)']).setValue(lastName);
    sheet.getRange(lastRowNum, headersJSON['Email']).setValue(email);
    sheet.getRange(lastRowNum, headersJSON['Téléphone']).setValue(phoneNo);
    sheet.getRange(lastRowNum, headersJSON['variant(from unbounce)']).setValue(variant);
    sheet.getRange(lastRowNum, headersJSON['page_uuid(from unbounce)']).setValue(page_uuid);
    sheet.getRange(lastRowNum, headersJSON['gclid(from unbounce)']).setValue(gclid);
    sheet.getRange(lastRowNum, headersJSON['javascript_prepopulated_value(from unbounce=timestamp)']).setValue(javascript_prepopulated_value);
    sheet.getRange(lastRowNum, headersJSON['status']).setValue(stage);
    sheet.getRange(lastRowNum, headersJSON['Timestamp 1']).setValue(timeStamp);
    
    //send mail
    sendMail({
      civilite: civilite,
      lastName: lastName,
      email: email
    });
  }
};

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
  
  return today;
};

function getTimeHhMmSs() {
  var now = new Date();
  var hh = now.getHours();
  var mm = now.getMinutes();
  var ss=now.getSeconds();
  
  var timenow = hh+':'+mm+':'+ss;
  
  return timenow;
};

function sendMail(JSONData) {
  var spreadsheet = SpreadsheetApp.openById(App.SPREADSHEET_ID),
      template = convertRangeToJSONArray(spreadsheet.getSheetByName(App.SHEETS.EMAIL_TEMPLATES).getDataRange().getValues()),
      configSheet = spreadsheet.getSheetByName(App.SHEETS.CONFIG);

  //Config values
  var spreadsheetLinkClient = configSheet.getRange("B6").getValue(),
      nom_client = configSheet.getRange("B2").getValue(),
      recipient1=configSheet.getRange("B3").getValue(),
      recipient2=configSheet.getRange("B4").getValue(),
      recipient3=configSheet.getRange("B5").getValue(),
      fromName=configSheet.getRange("B9").getValue();
  
  //Mail setup
  var recepients = template[1]['Recepients'].replace(/{ConfigSheet\/email_client1}/g, recipient1)
                                            .replace(/{ConfigSheet\/email_client2}/g, recipient2)
                                            .replace(/{ConfigSheet\/email_client3}/g, recipient3),
      subject = template[1]['Subject'].replace(/{civilite}/g, JSONData['civilite'])
                                      .replace(/{nom_de_famille}/g, JSONData['lastName'])
                                      .replace(/{ConfigSheet\/nom_client}/g, nom_client)
                                      .replace(/{email}/g, JSONData['email'])
                                      .replace(/{ConfigSheet\/spreadsheet_link_client}/g, spreadsheetLinkClient),
      message = template[1]['Message'].replace(/{civilite}/g, JSONData['civilite'])
                                      .replace(/{nom_de_famille}/g, JSONData['lastName'])
                                      .replace(/{ConfigSheet\/nom_client}/g, nom_client)
                                      .replace(/{email}/g, JSONData['email'])
                                      .replace(/{ConfigSheet\/spreadsheet_link_client}/g, spreadsheetLinkClient);
  
  GmailApp.sendEmail(recepients, subject, message, {
    name: fromName
  });
};
