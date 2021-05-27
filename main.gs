// SpreadSheet
// column C: preffix
// column D: operand
// column E: suffix
//
// example:
//   ",magazine,"
//   from:,no-reply@example.com,
//
// output example of sort rule:
//   {"magazine" from:no-reply@example.com}

const FUCK_ADDRESS_LIST_RANGE = "F2:F"

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Fuckin Mail Fucker');
  menu.addItem('振り分けルールを作成する', 'createSortRules');
  menu.addToUi();
}


function createSortRules() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(FUCK_ADDRESS_LIST_RANGE)
  var data = range.getValues();

  var result_roule = "{ ";

  for(var row = 0; row < range.getNumRows(); row++) {
    for(var col = 0; col < range.getNumRows(); col++) {
      v = data[row][col];

      if ( v === "" || v === null || v === undefined) {
        continue;
      }

      result_roule += v + " ";
    }
  }

  result_roule += "}";

  Logger.log(result_roule);
  Browser.msgBox(result_roule);

}
