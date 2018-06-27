function import(e) {
  // created by Sean Lowe
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName("Raw");
  var target = e;
  var range = source.getRange(3, 1, source.getLastRow()-2, 17).getValues();
  var arr = [];
  for (var i = 0; i < range.length; i++) {
    arr[i]=[];
    for (var j = 0; j < range[i].length; j++) {
      arr[i][j] = range[i][j];      // appointment time
      arr[i][j] = range[i][j];      // customer
      arr[i][j] = range[i][j];      // vehicle
      arr[i][j] = range[i][j];      // vin
      arr[i][j] = range[i][j];      // client advisor
    }
  }
  
  for (i = 0; i < arr.length-1; i++) {
    if (arr[i][3] == "" && arr[i][6] == "") { arr.splice(i, 1); i--; }
    if (arr[i][7] == arr[i+1][7]) { arr.splice(i+1, 1); i--; }
  }
  
  target.getRange(3, 1, arr.length, 17).setValues(arr);
}

function newSheet() {
  // created by Sean Lowe
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName('Master').copyTo(ss);
  SpreadsheetApp.flush();
  //var name = "test"; // uncomment this and the line under it for testing purposes
  //sheet.setName(name);
  var name = ui.prompt("Name of New Sheet", "Enter the date of the sheet you are creating (M/DD)", ui.ButtonSet.OK)
  sheet.setName(name.getResponseText());
  ss.setActiveSheet(sheet);
  import(sheet);
}