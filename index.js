function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{
    name: "Envía correos",
    functionName: "sendEmails"
  }];
  ss.addMenu("Mailing", menuEntries);
}

function sendEmails() {
  var Rainfall = DriveApp.getFilesByName("MACRO_GERENTE_2020")
  var Rainfall2 = DriveApp.getFilesByName("BASE CxP - REINTEGROS FEBRERO 13 OFICINA 19.xlsx")


  // Spreadsheets must be the first two ones in the right order
  var tagsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];

  var dataRange = tagsSheet.getRange(4, 2, 1, 1); // B4
  var mailColumn = dataRange.getValues()[0][0];

  dataRange = tagsSheet.getRange(7, 2, 1, 1); // B7
  var dataColumns = dataRange.getValues()[0][0];

  dataRange = tagsSheet.getRange(10, 2, 1, 1); // B10
  var subject = dataRange.getValues()[0][0];

  dataRange = tagsSheet.getRange(13, 2, 1, 1); // B13
  var content = dataRange.getValues()[0][0];




  // Fetch the range of operative cells A1:getLastColumn()getLastRow() and get the values
  dataRange = dataSheet.getRange(1, 1, dataSheet.getLastRow(), dataSheet.getLastColumn());
  var data = dataRange.getValues();

  // Loop over the rows
  for (i in data) {
    var row = data[i];
    var emailAddress = row[mailColumn.charCodeAt(0) - "A".charCodeAt(0)];

    // Apply only to valid Email addresses
    if (emailAddress.match("@") == "@") {
      var message = content;
      // Loop over the data tags
      for (var j = 0; j < dataColumns.length; j++) {
        var tag = "<" + dataColumns[j] + ">";
        if (message.match(tag) == tag) {

          var newText = row[dataColumns[j].charCodeAt(0) - "A".charCodeAt(0)];
          message = message.replace(tag, newText);


        }
      }

      var imagenes = "<img src=https://drive.google.com/uc?export=view&id=1AAyrJwp6lcbvE8MTYjx-TbG1XNCMfp9R><br><br>" + subject + "<br><br> " + message + "";
      MailApp.sendEmail(emailAddress, subject, message, {
        htmlBody: imagenes
      });


    }
  }
  Browser.msgBox("Mensajes enviados");
}