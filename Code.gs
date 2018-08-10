function import(target) {
  // created by Sean Lowe, 6/27/18
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //target = ss.getSheetByName("testSheet"); // uncomment this line for testing purposes
  var source = ss.getSheetByName("Raw");
  var range = source.getRange(1, 1, source.getLastRow(), 17).getValues();
  Logger.log(range);
  var arr = [];
  var count = 0;
  var check = false;
  var str;
  for (var i = 0; i < range.length; i++) {
    if ((range[i] == undefined || range[i][0] == "") 
        && (range[i+1] == undefined || range[i+1][0] == "") 
        && (range[i+2] == undefined || range[i+2][0] == "") 
        && (range[i+3] == undefined || range[i+3][0] == "")) { break; }
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
  range = ss.getSheetByName("List").getRange(2, 26, ss.getSheetByName("List").getLastRow()-2).getValues();
  for (i = 0; i < arr.length-1; i++) {
    check = true;
    str = arr[i][6].split("/");
    for (j = 0; j < range.length; j++) {
      if ( str[0] == range[j][0] ) { arr.splice(i, 1); i--; check = false; } // remove model-years on 'list'
    }
    if (check) {
      if (str[1] != "BMW") { arr.splice(i, 1); i--; } // remove non-BMW appointments
      else if (arr[i+1] != undefined && arr[i][3] == arr[i+1][3]) { arr.splice(i+1, 1); i--; } // remove appointments whose names are the same
      else if (arr[i][7] == "") { arr.splice(i, 1); i--; } // remove entries that have a blank VIN
      else if ((arr[i+1] != undefined && arr[i][7] == arr[i+1][7]) || arr[i+1][7] == "" ) { arr.splice(i+1, 1); i--; } // remove duplicate VIN appointments
    }
  }
  target.getRange(2, 1, arr.length, 17).setValues(arr);
}

function newSheet() {
  // created by Sean Lowe, 6/27/18
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet; var name; var check = false;
  while (!check) {
    name = ui.prompt("Name of New Sheet", "Enter the date of the sheet you are creating (M/DD)", ui.ButtonSet.OK_CANCEL);
    if (name.getSelectedButton() == ui.Button.OK) {
      name = name.getResponseText();
      name = name.replace("-","/"); name = name.replace("-","/");//These are DIFFERENT. Leave BOTH!
      if (name.split("/").length != 2) { ui.alert('Error', 'Please enter the date in the format M/DD', ui.ButtonSet.OK); }
      else if (parseInt(name.split("/")[0]) < 1 || parseInt(name.split("/")[0]) > 12){ ui.alert('Error', 'Please enter a valid month (1-12).', ui.ButtonSet.OK); }
      else if (name.split("/")[0].length > 1 && parseInt(name.split("/")[0]) != 12){ ui.alert('Error', 'Please enter the month in the format M/DD. Do not include a leading zero.', ui.ButtonSet.OK); }
      else if (parseInt(name.split("/")[1]) < 1 || parseInt(name.split("/")[1]) > 31){ ui.alert('Error', 'Please enter a valid day (1-31).', ui.ButtonSet.OK); }
      else if (ss.getSheetByName(name)!=null){
        sheet = ui.alert('Error', 'The sheet "'+name+'" already exists. Would you like to override the old sheet?', ui.ButtonSet.YES_NO_CANCEL);
        if (sheet == ui.Button.YES) { ss.deleteSheet(ss.getSheetByName(name)); }
        if (sheet == ui.Button.CANCEL) { ss.toast('Action cancelled. No sheets were created. No sheets were overridden.', 'Action cancelled'); return; }
      }
      else { check = true; }
    } else { ss.toast('Action cancelled. No sheets were created.', 'Action cancelled'); return; }
  }
  //var name = "test"; // uncomment this and the line under it for testing purposes
  //sheet.setName(name);
  ss.getSheetByName('Master').copyTo(ss).setName(name);
  sheet = ss.getSheetByName(name);
  ss.setActiveSheet(sheet);
  sheet.getRange(2, 28).setValue("=summarize(Q2:W,\"" + name + "\")");
  import(sheet);
  formUpdate();
  soldUpdate();
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
  var mpSold = 0;
  var esc = 0;
  var escSold = 0;
  var car = 0;
  var carSold = 0;
  
  var range = source.getRange(2, 17, source.getLastRow()-1, source.getLastColumn()-17).getValues();
  for (var i = 0; i < range.length; i++) {
    if ((range[i] == undefined || range[i][0] == "") && (range[i+1] == undefined || range[i+1][0] == "") && (range[i+2] == undefined || range[i+2][0] == "") && (range[i+3] == undefined || range[i+3][0] == "") ) { break; }
    if (range[i][6].toLowerCase() == "yes") { // update printed
     updated++;
      //Logger.log("reached updated increment");
      if (range[i][3].toLowerCase() == "yes") { // contacted
        contacted++;
        //Logger.log("reached contacted increment");
        if (range[i][4].toLowerCase() != "" && range[i][4].toLowerCase() != "n/a") {
          //turns++;
          var current = range[i][4].toUpperCase(); // turn
          var status = range[i][5].toUpperCase();  // status
          Logger.log(current);
          if (current == "MP & ESC") {
            if (status == "SOLD") { mpSold++; escSold++; }
            else { mp++; esc++; }
          } else if (current == "MP") {
            if (status == "SOLD") { mpSold++; }
            else { mp++; }
          } else if (current == "ESC") {
            if (status == "SOLD") { escSold++; }
            else { esc++; }
          } else if (current == "CAR") {
            if (status == "SOLD") { carSold++; }
            else { car++; }
          }
        }
      }
    }
  }
  return [[updated],[contacted],[mp],[esc],[car], [mpSold], [escSold], [carSold]];
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
  soldUpdate();
}

function refresh(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("calc").getRange("F16").setValue(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("calc").getRange("F16").getValue()+1);
}

//                                        0  1  2  3  4  5  6   7  8  9
//                                        |  X  X  C  T  S  UP  IP N  CA
// a b c d e f g h i j  k  l  m  n  o  p  q  r  s  t  u  v   w  x  y  z  aa ab
// 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22  23 24 25 26 27 28