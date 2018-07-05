function cas1(x) {
  //sheetName = "7/01";
  //CA = "Ben Wegener";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var list = []; //Array will be ordered: [CA,MP,ESC,CAR]
  var check = false;
  var CA;
  var min = 0;
  var max = 11;
  for (var k = min; k < max && k < sheets.length; k++) {
    var sheetName = sheets[k].getSheetName().toLowerCase();
    //Logger.log(sheetName);
    if (sheetName != "summary" && sheetName != "master" && sheetName != "raw" && sheetName != "list" && sheetName != "calc") {
      sheetName = sheets[k].getSheetName();
      //Logger.log(sheetName);
      var source = ss.getSheetByName(sheetName);
      var range = source.getRange(2, 17, source.getLastRow()-1, source.getLastColumn()-17).getValues();
      for (var j = 0; j < range.length; j++) {
        check = false;
        if ((range[j] == undefined || range[j][0] == undefined || range[j][0] == "") && (range[j+1] == undefined || range[j+1][0] == "") && (range[j+2] == undefined || range[j+2][0] == "")) { break; }
        if (range[j][9] != "" && range[j][9] != undefined){ CA = range[j][9].toUpperCase(); }
        else { CA = undefined; }
        var current = range[j][4].toUpperCase();
        if (CA != "" && CA != undefined) {
          for (var i = 0; i < list.length && !check; i++){
            if(CA == list[i][0]) {
              check=true;
              if (current == "MP & ESC") {
                list[i][1]++; //MP
                list[i][2]++; //ESC
              } else if (current == "MP") {
                list[i][1]++; //MP
              } else if (current == "ESC") {
                list[i][2]++; //ESC
              } else if (current == "CAR") {
                list[i][3]++; //CAR
              }
            }
          }
          if (!check) {
            list[list.length] = [CA,0,0,0];
            if (current == "MP & ESC") {
              list[list.length-1][1]++; //MP
              list[list.length-1][2]++; //ESC
            } else if (current == "MP") {
              list[list.length-1][1]++; //MP
            } else if (current == "ESC") {
              list[list.length-1][2]++; //ESC
            } else if (current == "CAR") {
              list[list.length-1][3]++; //CAR
            }
          }
        }
      }
    }
  }
  //Logger.log(list);
  if (list.length != 0) { return list; }
  else { return; }
}
function cas2(x) {
  //sheetName = "7/01";
  //CA = "Ben Wegener";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var list = []; //Array will be ordered: [CA,MP,ESC,CAR]
  var check = false;
  var CA;
  var min = 11;
  var max = 21;
  for (var k = min; k < max && k < sheets.length; k++) {
    var sheetName = sheets[k].getSheetName().toLowerCase();
    //Logger.log(sheetName);
    if (sheetName != "summary" && sheetName != "master" && sheetName != "raw" && sheetName != "list" && sheetName != "calc") {
      sheetName = sheets[k].getSheetName();
      //Logger.log(sheetName);
      var source = ss.getSheetByName(sheetName);
      var range = source.getRange(2, 17, source.getLastRow()-1, source.getLastColumn()-17).getValues();
      for (var j = 0; j < range.length; j++) {
        check = false;
        if ((range[j] == undefined || range[j][0] == undefined || range[j][0] == "") && (range[j+1] == undefined || range[j+1][0] == "") && (range[j+2] == undefined || range[j+2][0] == "")) { break; }
        if (range[j][9] != "" && range[j][9] != undefined){ CA = range[j][9].toUpperCase(); }
        else { CA = undefined; }
        var current = range[j][4].toUpperCase();
        if (CA != "" && CA != undefined) {
          for (var i = 0; i < list.length && !check; i++){
            if(CA == list[i][0]) {
              check=true;
              if (current == "MP & ESC") {
                list[i][1]++; //MP
                list[i][2]++; //ESC
              } else if (current == "MP") {
                list[i][1]++; //MP
              } else if (current == "ESC") {
                list[i][2]++; //ESC
              } else if (current == "CAR") {
                list[i][3]++; //CAR
              }
            }
          }
          if (!check) {
            list[list.length] = [CA,0,0,0];
            if (current == "MP & ESC") {
              list[list.length-1][1]++; //MP
              list[list.length-1][2]++; //ESC
            } else if (current == "MP") {
              list[list.length-1][1]++; //MP
            } else if (current == "ESC") {
              list[list.length-1][2]++; //ESC
            } else if (current == "CAR") {
              list[list.length-1][3]++; //CAR
            }
          }
        }
      }
    }
  }
  //Logger.log(list);
  if (list.length != 0) { return list; }
  else { return; }
}
function cas3(x) {
  //sheetName = "7/01";
  //CA = "Ben Wegener";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var list = []; //Array will be ordered: [CA,MP,ESC,CAR]
  var check = false;
  var CA;
  var min = 21;
  var max = 31;
  for (var k = min; k < max && k < sheets.length; k++) {
    var sheetName = sheets[k].getSheetName().toLowerCase();
    //Logger.log(sheetName);
    if (sheetName != "summary" && sheetName != "master" && sheetName != "raw" && sheetName != "list" && sheetName != "calc") {
      sheetName = sheets[k].getSheetName();
      //Logger.log(sheetName);
      var source = ss.getSheetByName(sheetName);
      var range = source.getRange(2, 17, source.getLastRow()-1, source.getLastColumn()-17).getValues();
      for (var j = 0; j < range.length; j++) {
        check = false;
        if ((range[j] == undefined || range[j][0] == undefined || range[j][0] == "") && (range[j+1] == undefined || range[j+1][0] == "") && (range[j+2] == undefined || range[j+2][0] == "")) { break; }
        if (range[j][9] != "" && range[j][9] != undefined){ CA = range[j][9].toUpperCase(); }
        else { CA = undefined; }
        var current = range[j][4].toUpperCase();
        if (CA != "" && CA != undefined) {
          for (var i = 0; i < list.length && !check; i++){
            if(CA == list[i][0]) {
              check=true;
              if (current == "MP & ESC") {
                list[i][1]++; //MP
                list[i][2]++; //ESC
              } else if (current == "MP") {
                list[i][1]++; //MP
              } else if (current == "ESC") {
                list[i][2]++; //ESC
              } else if (current == "CAR") {
                list[i][3]++; //CAR
              }
            }
          }
          if (!check) {
            list[list.length] = [CA,0,0,0];
            if (current == "MP & ESC") {
              list[list.length-1][1]++; //MP
              list[list.length-1][2]++; //ESC
            } else if (current == "MP") {
              list[list.length-1][1]++; //MP
            } else if (current == "ESC") {
              list[list.length-1][2]++; //ESC
            } else if (current == "CAR") {
              list[list.length-1][3]++; //CAR
            }
          }
        }
      }
    }
  }
  //Logger.log(list);
  if (list.length != 0) { return list; }
  else { return; }
}
function cas4(x) {
  //sheetName = "7/01";
  //CA = "Ben Wegener";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var list = []; //Array will be ordered: [CA,MP,ESC,CAR]
  var check = false;
  var CA;
  var min = 31;
  var max = 41;
  for (var k = min; k < max && k < sheets.length; k++) {
    var sheetName = sheets[k].getSheetName().toLowerCase();
    //Logger.log(sheetName);
    if (sheetName != "summary" && sheetName != "master" && sheetName != "raw" && sheetName != "list" && sheetName != "calc") {
      sheetName = sheets[k].getSheetName();
      Logger.log(sheetName);
      var source = ss.getSheetByName(sheetName);
      var range = source.getRange(2, 17, source.getLastRow()-1, source.getLastColumn()-17).getValues();
      for (var j = 0; j < range.length; j++) {
        check = false;
        if ((range[j] == undefined || range[j][0] == undefined || range[j][0] == "") && (range[j+1] == undefined || range[j+1][0] == "") && (range[j+2] == undefined || range[j+2][0] == "")) { break; }
        if (range[j][9] != "" && range[j][9] != undefined){ CA = range[j][9].toUpperCase(); }
        else { CA = undefined; }
        var current = range[j][4].toUpperCase();
        if (CA != "" && CA != undefined) {
          for (var i = 0; i < list.length && !check; i++){
            if(CA == list[i][0]) {
              check=true;
              if (current == "MP & ESC") {
                list[i][1]++; //MP
                list[i][2]++; //ESC
              } else if (current == "MP") {
                list[i][1]++; //MP
              } else if (current == "ESC") {
                list[i][2]++; //ESC
              } else if (current == "CAR") {
                list[i][3]++; //CAR
              }
            }
          }
          if (!check) {
            list[list.length] = [CA,0,0,0];
            if (current == "MP & ESC") {
              list[list.length-1][1]++; //MP
              list[list.length-1][2]++; //ESC
            } else if (current == "MP") {
              list[list.length-1][1]++; //MP
            } else if (current == "ESC") {
              list[list.length-1][2]++; //ESC
            } else if (current == "CAR") {
              list[list.length-1][3]++; //CAR
            }
          }
        }
      }
    }
  }
  //Logger.log(list);
  if (list.length != 0) { return list; }
  else { return; }
}
function calc(a,b,c,d) {
  var all = [a,b,c,d];
  var final = [];
  var check = false;
  for (var i = 0; i < all.length; i++) {
    if (all[i] != null && all[i] != undefined) {
      for (var j = 0; j < all[i].length && !check; j++) {
        check = false;
        if (all[i][j][0] != "" && all[i][j][0] != undefined) {
          for (var k = 0; k < final.length; k++) {
            if (all[i][j][0] == final[k][0]){
              check = true;
              final[k][1] += all[i][j][1];
              final[k][2] += all[i][j][2];
              final[k][3] += all[i][j][3];
            }
          }
          if (!check) {
            final[final.length] = all[i][j];
          }
        }
      }
    }
  }
  return final;
}

