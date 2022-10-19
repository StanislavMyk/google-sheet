/** @OnlyCurrentDoc */
function onOpen() {
  // Logger.log('sdd')
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Worklist');
  // let da = new Date();
 
  ss.getRange('D1').setFormula(`sumif($F$6:$F$1000;"="&TODAY();$L$6:$L$1000)`);
  ss.getRange('D3').setFormula(`SUMIFS(L6:L1000;F6:F1000;">="& EOMONTH(TODAY();-1)+1;F6:F1000;"<="& EOMONTH(TODAY();0))`);
  // week range
  var curr = new Date; // get current date
  var first = curr.getDate() - curr.getDay(); // First day is the day of the month - the day of the week
  var last = first + 6; // last day is the first day + 6

  var firstday = new Date(curr.setDate(first));
  var lastday = new Date(curr.setDate(last));
  ss.getRange('D2').setFormula(`sumifs(L6:L1000;F6:F1000; ">="&DATE(${firstday.getFullYear()};${firstday.getMonth()+1};${firstday.getDate()}); F6:F1000;"<="&DATE(${lastday.getFullYear()};${lastday.getMonth()+1};${lastday.getDate()}))`);
}

function startProject() {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Worklist');
  Logger.log(ss.getRange('B5').getValue())
  if (ss.getRange('B5').getValue() === '') {
    SpreadsheetApp.getUi().alert('Please select project id.');
    return;
  }
  let date = new Date();
  // Logger.log(date);
  ss.getRange('F5').setValue(`${date.getDate()}-${date.getMonth() + 1}-${date.getFullYear()}`);
  ss.getRange('G5').setValue(`${date.getHours()}:${date.getMinutes()}:${date.getSeconds()}`);
}
function stopProject() {

  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Worklist');
  if (ss.getRange('G5').getValue() === '') {
    SpreadsheetApp.getUi().alert('Please start project.');
    return;
  }
  let date = new Date();
  ss.getRange('H5').setValue(`${date.getHours()}:${date.getMinutes()}:${date.getSeconds()}`);
  // set formulas
  ss.getRange('I5').setFormula('=(HOUR(H5-G5)*60) + MINUTE(H5-G5)');
  ss.getRange('J5').setFormula('=MROUND(I5;15)');
  ss.getRange('L5').setFormula('=J5*K5');

  // add new work line
  ss.getRange('B5:M5').insertCells(SpreadsheetApp.Dimension.ROWS);

  // total count builder
  ss.getRange('D1').setFormula(`sumif($F$6:$F$1000;"="&TODAY();$L$6:$L$1000)`);
  ss.getRange('D3').setFormula(`SUMIFS(L6:L1000;F6:F1000;">="& EOMONTH(TODAY();-1)+1;F6:F1000;"<="& EOMONTH(TODAY();0))`);
  // week range
  var curr = new Date; // get current date
  var first = curr.getDate() - curr.getDay(); // First day is the day of the month - the day of the week
  var last = first + 6; // last day is the first day + 6

  var firstday = new Date(curr.setDate(first));
  var lastday = new Date(curr.setDate(last));
  ss.getRange('D2').setFormula(`sumifs(L6:L1000;F6:F1000; ">="&DATE(${firstday.getFullYear()};${firstday.getMonth()+1};${firstday.getDate()}); F6:F1000;"<="&DATE(${lastday.getFullYear()};${lastday.getMonth()+1};${lastday.getDate()}))`);

  // spent time builder on project sheet
  let ps=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  ps.getRange('L2:L1000').setFormula('sumif(Worklist!$B$6:$B$1000; "="&B2; Worklist!$I$6:$I$1000)')
}

// project selector trigger on worklist

function onEdit(e) {
  // Logger.log('this is trigger')
  // id selector on worklist
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let cell = ss.getActiveCell();
  if (cell.getA1Notation() === 'B5' && ss.getSheetName() === 'Worklist') {
    let ps = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
    let spn = Number(cell.getValue());
    let projectList = ps.getRange('B2:O1000').getValues();
    let i=0;
    for (row of projectList) {
      if (Number(row[0]) === spn) {
        ss.getRange('C5').setValue(row[1]);
        ss.getRange('D5').setValue(row[2]);
        ss.getRange('K5').setValue(row[3]);
        ss.getRange('M5').insertCheckboxes();
        ss.getRange('M5').setFormula(`and(Projects!I${i+2};true)`);
        break;
      }
      i++;
    }

  }

  // completed function process in projectlist 

}

function insertNewProject() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  spreadsheet.getRange('B2:O2').activate();
  spreadsheet.getRange('B2:O2').insertCells(SpreadsheetApp.Dimension.ROWS);
  var cell = spreadsheet.getRange('B2');
  cell.setValue(Number(spreadsheet.getRange('B3').getValue()) + 1)
  var completeCell = spreadsheet.getRange('I2');
  var paidCell = spreadsheet.getRange('J2');
  completeCell.insertCheckboxes();
  paidCell.insertCheckboxes();
  var deadlineCell = spreadsheet.getRange('H2');
  deadlineCell.setValue(new Date());
  var rateCell = spreadsheet.getRange('E2');
  rateCell.setValue('€1');
  var timespentCell=spreadsheet.getRange('L2');
  timespentCell.setFormula('sumif(Worklist!$B$6:$B$1000; "="&B2; Worklist!$I$6:$I$1000)');

  // set id selector on worklist
  var spreadsheet1 = SpreadsheetApp.getActive();
  spreadsheet1.getRange('B2:B55').activate();
  spreadsheet1.getRange('Worklist!B5').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .setHelpText('Click and enter a value from range Projects!B2:B55')
    .requireValueInRange(spreadsheet.getRange('Projects!$B$2:$B$1000'), true)
    .build());
};
function filterMangaer() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  let range = spreadsheet.getRange('B1:O1000');
  let filter = spreadsheet.getFilter();
  if (filter === null) {
    range.createFilter();
  }
  return range.getFilter();
}
function hideCompleted() {

  let filter = filterMangaer();
  let criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues([true, ''])
    .build();
  filter.setColumnFilterCriteria(9, criteria);
}

function showAll() {
  let filter = filterMangaer();
  let criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues([''])
    .build();
  filter.setColumnFilterCriteria(9, criteria);
  filter.remove();
}

function showCompleted() {
  let filter = filterMangaer();
  let criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues([false, ''])
    .build();
  filter.setColumnFilterCriteria(9, criteria);
}

function filterMangaer_worklist(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Worklist');
  let range = spreadsheet.getRange('B4:M1000');
  let filter = spreadsheet.getFilter();
  if (filter === null) {
    range.createFilter();
  }
  return range.getFilter();
}
function hideCompleted_worklist() {

  let filter = filterMangaer_worklist();
  let criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues([true, ''])
    .build();
  filter.setColumnFilterCriteria(13, criteria);
}

function showAll_worklist() {
  let filter = filterMangaer_worklist();
  let criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues([''])
    .build();
  filter.setColumnFilterCriteria(13, criteria);
  filter.remove();
}

function showCompleted_worklist() {
  let filter = filterMangaer_worklist();
  let criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues([false, ''])
    .build();
  filter.setColumnFilterCriteria(13, criteria);
}
