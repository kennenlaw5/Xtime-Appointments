function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Utilities').addSubMenu(ui.createMenu('Contact Kennen').addItem('By Phone','phoneKennen')
                                        .addItem('By Email','emailKennen')).addItem('Create New Sheet', 'newSheet').addToUi();
                                        //.addItem('Summarize Spreadsheet', 'summarize').addToUi();
  ss.getSheetByName("Master").hideSheet();
}

function phoneKennen() {
  SpreadsheetApp.getUi().alert('Call or text (720) 317-5427');
}

function emailKennen() {
  // Created By Kennen Lawrence
  var ui = SpreadsheetApp.getUi();
  var input = ui.prompt('Email Sheet Creator','Describe the issue you\'re having in the box below, then press "Ok" to submit your issue via email:',ui.ButtonSet.OK_CANCEL);
  if (input.getSelectedButton() == ui.Button.OK) {
    MailApp.sendEmail('kennen.lawrence@a2zsync.com','HELP Xtime Appointments',input.getResponseText()+"\n\n\nhttps://docs.google.com/spreadsheets/d/1xdS_MC3ZSGwZENtMQAHRETeWlE8pIY3zaCPOS1xgwHs/edit#gid=1629835607");
    SpreadsheetApp.getActiveSpreadsheet().toast('Email sent successfully! We will get back to you as quick as possible!', 'Success!')
  } else if (input.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log('User cancelled');
  }
}

function formUpdate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet = ss.getSheetByName("Summary");
  var formulas = sheet.getRange(5, 5, 1, 3).getFormulas();
  var updated = []; var first = true; var current;
  for (var i = 0; i < formulas[0].length; i++) {
    updated[i] = "=SUM(";
    first = true;
    for (var j = 0; j < sheets.length; j++) {
      current = sheets[j].getSheetName().toLowerCase();
      if (current != "summary" && current != "master" && current != "raw" && current != "list") {
        if (first) { updated[i] += "'" + sheets[j].getSheetName() + "'!$AB" + (i+1); first=false; }
        else { updated[i] += ",'" + sheets[j].getSheetName() + "'!$AB" + (i+1); }
        if (j+1 >= sheets.length) { updated[i] += ")"; }
      }
    }
  }
  if (sheets.length == 4) { updated[0] += ")"; updated[1] += ")"; updated[2] += ")"; }
  sheet.getRange(5, 5, 1, 3).setValues([updated]);
}