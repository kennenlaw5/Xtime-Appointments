function import(target) {
  // created by Sean Lowe, 6/27/18
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName("Raw");
  var range = source.getRange(3, 1, source.getLastRow()-2, 17).getValues();
  var arr = [];
  var count = 0;
  for (var i = 0; i < range.length; i++) {
    if (range[i][0]=="" && range[i+1][0] == "" && range[i+2][0] == "") { i = range.length-1; }
    else {
      if (range[i][0] != 0 && range[i][0] != "Confirmation Key" && range[i][0] != "") {
        arr[count]=[];
        for (var j = 0; j < range[i].length; j++) {
          arr[count][j] = range[i][j];
        }
        count++;
      }
    }
  }
  for (i = 0; i < arr.length-1; i++) {
    if (arr[i][7] == arr[i+1][7]) { arr.splice(i+1, 1); i--; }
  }
  target.getRange(2, 1, arr.length, 17).setValues(arr);
}

function newSheet() {
  // created by Sean Lowe, /6/27/18
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet;
  var name = ui.prompt("Name of New Sheet", "Enter the date of the sheet you are creating (M/DD)", ui.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() == ui.Button.OK) {
    //var name = "test"; // uncomment this and the line under it for testing purposes
    //sheet.setName(name);
    ss.getSheetByName('Master').copyTo(ss).setName(name.getResponseText());
    sheet = ss.getSheetByName(name.getResponseText());
    ss.setActiveSheet(sheet);
    import(sheet);
  }
}

function summarize() {
  // created by Sean Lowe, 6/28/18
  Logger.log("reached summarize");
}