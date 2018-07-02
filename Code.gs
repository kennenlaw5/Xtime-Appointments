function import(target) {
  // created by Sean Lowe, 6/27/18
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName("Raw");
  var range = source.getRange(1, 1, source.getLastRow(), 17).getValues();
  var arr = [];
  var count = 0;
  for (var i = 0; i < range.length; i++) {
    if (range[i][0]=="" && range[i+1][0] == "" && range[i+2][0] == "") { break; }
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
  // created by Sean Lowe, 6/27/18
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet, formulas;
  var name = ui.prompt("Name of New Sheet", "Enter the date of the sheet you are creating (M/DD)", ui.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() == ui.Button.OK) {
    //var name = "test"; // uncomment this and the line under it for testing purposes
    //sheet.setName(name);
    ss.getSheetByName('Master').copyTo(ss).setName(name.getResponseText());
    sheet = ss.getSheetByName(name.getResponseText());
    ss.setActiveSheet(sheet);
    sheet.getRange(1, 28).setValue("=summarize(Q2:W,\"" + name.getResponseText() + "\")");
    formulas = ss.getSheetByName("Summary").getRange(14, 5, 1, 3).getFormulas();
    for (var i = 0; i < formulas[0].length; i++) {
      formulas[0][i] = formulas[0][i].split(")")[0] + ",'" + name.getResponseText() + "'!$AB" + (i+1) + ")";
    }
    ss.getSheetByName("Summary").getRange(14, 5, 1, 3).setValues(formulas);
    import(sheet);
  }
}

function summarize(x, sheetName) {
  // created by Sean Lowe, 6/29/18
  //Logger.log("reached summarize");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = null;
  var source = ss.getSheetByName(sheetName);
  var updated = 0;
  var contacted = 0;
  var turns = 0;
  
  var range = source.getRange(2, 17, source.getLastRow()-1, source.getLastColumn()-17).getValues();
  for (var i = 0; i < range.length; i++) {
    if (range[i][0] == "" && range[i+1][0] == "" && range[i+2][0] == "") { break; }
    if (range[i][6].toLowerCase() == "yes") {
     updated++;
      //Logger.log("reached updated increment");
      if (range[i][3].toLowerCase() == "yes") {
        contacted++;
        //Logger.log("reached contacted increment");
        if (range[i][4].toLowerCase() != "" && range[i][4].toLowerCase() != "n/a") {
          turns++;
        }
      }
    }
  }
  return [[updated],[contacted],[turns]];
  //Logger.log("updated = " + updated);
  //Logger.log("contacted = " + contacted);
  //Logger.log("turns = " + turns);
  //Logger.log("i = " + i);
}

function newMonth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var name;
  for (var i = 0; i < sheets.length; i++) {
    name = sheets[i].getName().toLowerCase();
    if (name != "summary" || name != "master" || name != "raw") {
      ss.deleteSheet(sheets[i]);
    }
  }
}

//                                        0  1  2  3  4  5  6  7
//                                        |  X  X  C  T  S  UP IP
// a b c d e f g h i j  k  l  m  n  o  p  q  r  s  t  u  v   w  x  y  z  aa ab
// 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22  23 24 25 26 27 28