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




/*
What will this have to do?

Create a copy of the 'Master' sheet and rename it to whatever the day is
    done
hide columns A, C, E-F, J - P, R - S
    done
pull data from imported sheet (columns 1, 3, 6, 7, 16)
push data to new sheet (columns 1, 3, 6, 7, 16) <-- yes they are the same
add in the 
'contact,' (column 19)
'turn,'    (column 20)
'status,'  (column 21)
'update printed,' (column 22)
'in person,'      (column 23)
'notes,'   (column 24)
'CA'       (column 25)
columns to end of new sheet
*/