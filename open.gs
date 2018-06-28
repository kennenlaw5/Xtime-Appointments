function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Utilities').addSubMenu(ui.createMenu('Contact Kennen').addItem('By Phone','phoneKennen')
                                        .addItem('By Email','emailKennen')).addItem('Create New Sheet', 'newSheet')
                                        .addItem('Summarize Spreadsheet', 'summarize').addToUi();
}

function phoneKennen() {
  SpreadsheetApp.getUi().alert('Call or text (720) 317-5427');
}

function emailKennen() {
  // Created By Kennen Lawrence
  var ui = SpreadsheetApp.getUi();
  var input = ui.prompt('Email Sheet Creator','Describe the issue you\'re having in the box below, then press "Ok" to submit your issue via email:',ui.ButtonSet.OK_CANCEL);
  if (input.getSelectedButton() == ui.Button.OK) {
    MailApp.sendEmail('kennen.lawrence@schomp.com','HELP Auto Notification',input.getResponseText());
  } else if (input.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log('User cancelled');
  }
}