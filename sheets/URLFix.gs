// onOpen: Add our menu items.
function onOpen() {
  SpreadsheetApp.getActive()
  .addMenu("URLFix",[
    {name: "Fix URL's in selection", functionName: "fixURLs_"}
    ]);
}

// fixURLs_: Iterate through the user selected range of cells
// and attempt to find the HYPERLINK filename in Google Drive
// then update the cell with the drive link.

// TODO: Error handling for non URL's
// TODO: Allow just strings to work instead of formulas
// TODO: Report success/failure counters
// TODO: Break this function up, it's gross.
// TODO: Optimize order of operations for minimal compute time
function fixURLs_() {
  var fixed = 0;
  var notFound = 0;
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  if (range == null) {
    ui.alert(
      'Select a Column or Cell',
      "Select the column or cells you want to fix URL's for.",
      ui.ButtonSet.OK);
  }
  
  var data = range.getFormulas();
  
  var re = new RegExp("=hyperlink\\\(\"([^\"]+)\",(\s+)?\"([^\"]+)", "gi");
  
  // Loop through all values and attempt to fix all URL's.
  for (var y = 0; y < data.length; y++){
    for (var x = 0; x < data[y].length; x++){
      formula = data[y][x];
      if (formula.length <= 0){
        continue;
      }
      res = re.exec(formula);
      if (res == null) {
        notFound++;
        continue; // TODO: Invalid cell, warn?
      }
      fileName = res[1];
      fileList = DriveApp.getFilesByName(fileName);
      // TODO: Handle file not found counter
      // TODO: What do we do if we find multiple files with the same name? Do we prompt?
      if (!fileList.hasNext()) {
        notFound++;
      }
      while (fileList.hasNext()){
        var file = fileList.next();
        var driveUrl = file.getUrl();
        range.getCell(y+1, x+1).setFormula('=HYPERLINK("' + driveUrl + '", "' + res[3] + '")')
        fixed++;
      }
    }
  }
  ui.alert(
    "URL's Fixed",
    "Fixed " + fixed + " URL's and couldn't find " + notFound + " documents in your drive.",
    ui.ButtonSet.OK);
}
