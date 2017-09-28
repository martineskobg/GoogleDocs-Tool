function copySheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      acSheet = ss.getActiveSheet(),
      nameSheet = acSheet.getRange(2, 3).getValue(),
      oldNameSheet = acSheet.getRange(2, 4).getValue(),
      sheet = ss.getSheetByName(oldNameSheet),
      range = ss.getActiveRange(),
      val = range.getValues(),
      len = val.length,
      index = range.getRowIndex(),
      end = index + (len - 1);
  
      range = ss.getRange('A' + index + ':A' + end);
  var values = range.getValues();
           
  for(var row = 0;row < range.getHeight(); row++){
      var dest = SpreadsheetApp.openById(values[row][0]);
      SpreadsheetApp.setActiveSpreadsheet(dest);
      var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet);
      if(s != null) continue ;

      SpreadsheetApp.setActiveSheet(dest.getSheetByName(oldNameSheet));
       dest.renameActiveSheet(nameSheet);
       dest.duplicateActiveSheet();
       dest.renameActiveSheet(oldNameSheet);
         
      }
}

function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      menuEntries = [];
  // When the user clicks on "addMenuExample" then "Menu Entry 1", the function function1 is
  // executed.

  menuEntries.push({
    name: 'Duplicate Sheet',
    functionName: 'copySheet'
  });
  menuEntries.push(null); // line separator
  ss.addMenu('Избери функция', menuEntries);
}
onOpen();