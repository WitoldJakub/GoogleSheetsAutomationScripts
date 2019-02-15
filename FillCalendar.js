function fillCalendar() {
  // Accesing the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Reading the holiday table relevant data
  var names_list = sheet.getRange("B5:B105").getValues();
  // Reading start and end of holidays dates
  // (hidden colums in the sheet - integer values)
  var date_leave = sheet.getRange("I5:I105").getValues();
  var date_back = sheet.getRange("J5:J105").getValues();
  // Output values of the status
  var status = sheet.getRange("H5:H105").getValues();

  // Coordinates of top-left cell of the calendar
  var r1 = 6; // row 6
  var c1 = 15; // column "O"

  var date = sheet.getRange("O3:NP3").getValues();
  //The row of dates is read as 1x365 array, the trasposition to 365x1 is neede
  date = date[0].map(function (_, c) { return date.map(function (r) { return r[c]; }); });
  Logger.log(date.length);
  var names = sheet.getRange("L6:L17").getValues()
  Logger.log(names);

// Clear calendar
     sheet.getRange(r1, c1, r1 + 11, c1 + 366).setValue("");

// Fill calendar
   for (var i = 0; i < names.length; i++) {
    var nameI = String(names[i]);
    for (var j = 0; j < date.length; j++){
      for (var n = 0; n < names_list.length; n++){
       var nameN = String(names_list[n]);
       if ( date_leave[n] <= date[j] && date_back[n] >= date[j] && nameN.equals(nameI) )
        { sheet.getRange(i + r1, j + c1).setValue(status[n]);
    }
   }
  }
 }
}
