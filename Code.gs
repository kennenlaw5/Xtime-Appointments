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
    str = arr[i][6].split("/");
    if ( (str[0] == "2017" || str[0] == "2018" || str[0] == "2019") || str[1] != "BMW") { arr.splice(i, 1); i--; }
    if (arr[i][7] == arr[i+1][7]) { arr.splice(i+1, 1); i--; }
  }
  target.getRange(2, 1, arr.length, 17).setValues(arr);
}

function newSheet() {
  // created by Sean Lowe, 6/27/18
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
    sheet.getRange(2, 28).setValue("=summarize(Q2:W,\"" + name.getResponseText() + "\")");
    formUpdate();
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
  //var turns = 0;
  var mp = 0;
  var esc = 0;
  var car = 0;
  
  var range = source.getRange(2, 17, source.getLastRow()-1, source.getLastColumn()-17).getValues();
  for (var i = 0; i < range.length; i++) {
    if ((range[i] == undefined || range[i][0] == "") && (range[i+1] == undefined || range[i+1][0] == "") && (range[i+2] == undefined || range[i+2][0] == "")) { break; }
    if (range[i][6].toLowerCase() == "yes") {
     updated++;
      //Logger.log("reached updated increment");
      if (range[i][3].toLowerCase() == "yes") {
        contacted++;
        //Logger.log("reached contacted increment");
        if (range[i][4].toLowerCase() != "" && range[i][4].toLowerCase() != "n/a") {
          //turns++;
          var current = range[i][4].toUpperCase();
          Logger.log(current);
          if (current == "MP & ESC") {
            mp++; esc++;
          } else if (current == "MP") {
            mp++;
          } else if (current == "ESC") {
            esc++;
          } else if (current == "CAR") {
            car++;
          }
        }
      }
    }
  }
  return [[updated],[contacted],[mp],[esc],[car]];
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
    //Logger.log("current sheet name is " + name);
    if (name != "summary" && name != "master" && name != "raw" && name != "list" && name != "calc") {          
      //Logger.log("able to delete current sheet: " + name);
      ss.deleteSheet(sheets[i]);
    }
  }
  formUpdate();
}

function refresh(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("calc").getRange("F16").setValue(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("calc").getRange("F16").getValue()+1);
}

//                                        0  1  2  3  4  5  6   7  8  9
//                                        |  X  X  C  T  S  UP  IP N  CA
// a b c d e f g h i j  k  l  m  n  o  p  q  r  s  t  u  v   w  x  y  z  aa ab
// 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22  23 24 25 26 27 28