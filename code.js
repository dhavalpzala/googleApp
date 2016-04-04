var App = {
  SPREADSHEET_ID: '1awUjgNSOrudnuyXpxI1ZjY05ousoPTT5R9_nuW3HwJ4',
  SHEETS: {
    TARGETS: 'Targets',
    SEQUENCE_A: 'SequenceA',
    SEQUENCE_B: 'SequenceB',
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

// function to send the first email only
function schedule() {
  var spreadsheet = SpreadsheetApp.openById(App.SPREADSHEET_ID),
      targetsSheet = spreadsheet.getSheetByName(App.SHEETS.TARGETS),
      targetsRange = targetsSheet.getDataRange().getValues(),
      targets = convertRangeToJSONArray(targetsRange),
      sequencesA = convertRangeToJSONArray(spreadsheet.getSheetByName(App.SHEETS.SEQUENCE_A).getDataRange().getValues()),
      sequencesB = convertRangeToJSONArray(spreadsheet.getSheetByName(App.SHEETS.SEQUENCE_B).getDataRange().getValues()),
      sequenceCountColIndex = GetObjectKeyIndex(targets[0], 'Sequence Count'),
      config = convertRangeToJSONArray(spreadsheet.getSheetByName(App.SHEETS.CONFIG).getDataRange().getValues()),
      alias = Session.getActiveUser().getEmail(),
      name = "";
  
  if(GmailApp.getAliases().indexOf(config[0]['From']) > -1) {
    alias = config[0]['From'];
    name = config[0]['FromName'];
  }
  
  for( var index = 0; index < targets.length; index++) {
    var sendDateTime = targets[index]['SendDateTime'],
        sequenceCount = targets[index]['Sequence Count'],
        active = targets[index]['Active'],
        sequence = targets[index]['Sequence'],
        sequences;
      
    if(sequence === 'A') {
      sequences = sequencesA;
    } 
    else if(sequence === 'B') {
      sequences = sequencesB;
    }
    
    if(sequenceCount) {
      //find next sequence datetime
      var delayTime = sequences[sequenceCount]['Delay Time'];
      
      var delayTimeArray = delayTime.split(" "),
          delayDays = 0,
          delayHours = 0,
          delayMinutes = 0;
      
      delayTimeArray.forEach(function(time) {
        if(time.indexOf("d") > 0) {
          delayDays = parseInt(time);
        }
        else if(time.indexOf("h") > 0) {
          delayHours = parseInt(time);
        }
        else if(time.indexOf("m") > 0) {
          delayMinutes = parseInt(time);
        }
      });
      
      if(delayDays) {
        sendDateTime.setDate(sendDateTime.getDate() + delayDays);
      }
      
      if(delayHours) {
        sendDateTime.setHours(sendDateTime.getHours() + delayHours);
      }
      
      if(delayMinutes) {
        sendDateTime.setMinutes(sendDateTime.getMinutes() + delayMinutes);
      }
    } 
    else {
      sequenceCount = 0;
    }
    
    if(checkForSendingMail(sendDateTime, active)) {     
      //to send mail
      sendMail({
        civility: targets[index]['Civility'],
        firstName: targets[index]['FirstName'],
        lastName: targets[index]['LastName'],
        company: targets[index]['Company'],
        email: targets[index]['Email'],
        subject: sequences[sequenceCount]['Subject'],
        message: sequences[sequenceCount]['Message']
      }, alias, name);
      
      //increment sequece count
      targetsSheet.getRange(index + 2, sequenceCountColIndex).setValue(++sequenceCount);
    } 
  }
};

function checkForSendingMail(sendDateTime, active) {
  if(sendDateTime) {
    var currentTime = Date.now(),
        sendTime = new Date(sendDateTime).getTime(),
        expectedDiff = 3600000; // in milli seconds
    
    return active !== 'no' && currentTime >= sendTime && (currentTime - sendTime) < expectedDiff;
  }
  else {
    return false;
  }
};

function sendMail(mail, alias, name) {
  var recipient = mail.email,
      subject = mail.subject,
      body = mail.message.replace("{!Civility}", mail.civility)
                         .replace("{!FirstName}", mail.firstName)
                         .replace("{!LastName}", mail.lastName)
                         .replace("{!Company}", mail.company);
  
  GmailApp.sendEmail(recipient, subject, '', {
    htmlBody: body,
    from: alias,
    name: name
  });
};

function GetObjectKeyIndex(obj, keyToFind) {
  var index = 1, key;
  for (key in obj) {
    if (key == keyToFind) {
      return index;
    }
    index++;
  }
  return null;
};
