//Comfortable living while owning a car and not driving more than 5 miles a day on average at $0.50 per mile, as well as saving for 6-month backup
function config1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRange('H4').setValue(1);
  ss.getRange('H5').setValue(0);
  ss.getRange('H6').setValue(0.7732);
  ss.getRange('H7').setValue(0.056603987);
  ss.getRange('H8').setValue(0.3214);
  ss.getRange('H9').setValue(0.51925);
  ss.getRange('H10').setValue(0.0635);
  ss.getRange('H11').setValue(0);
  ss.getRange('H12').setValue(1);
  ss.getRange('H13').setValue(1);
  ss.getRange('H14').setValue(1);
}

//Going all out on all monthly expenses, including living savings and necessary savings
function config2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRange('H4').setValue(1);
  ss.getRange('H5').setValue(1);
  ss.getRange('H6').setValue(1);
  ss.getRange('H7').setValue(1);
  ss.getRange('H8').setValue(1);
  ss.getRange('H9').setValue(1);
  ss.getRange('H10').setValue(1);
  ss.getRange('H11').setValue(1);
  ss.getRange('H12').setValue(1);
  ss.getRange('H13').setValue(1);
  ss.getRange('H14').setValue(1);
}

//Living comfortably while saving up to buy a car
function config3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRange('H4').setValue(1);
  ss.getRange('H5').setValue(0);
  ss.getRange('H6').setValue(0);
  ss.getRange('H7').setValue(0.056603987);
  ss.getRange('H8').setValue(0.3214);
  ss.getRange('H9').setValue(0.51925);
  ss.getRange('H10').setValue(0.1889);
  ss.getRange('H11').setValue(0);
  ss.getRange('H12').setValue(1);
  ss.getRange('H13').setValue(1);
  ss.getRange('H14').setValue(1);
}

//Live comfortably while owning a car, not driving it for more than 5 miles/day at $0.50/mile for gas on average, and saving up for Aman and 6-month backup
function config4() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRange('H4').setValue(1);
  ss.getRange('H5').setValue(0);
  ss.getRange('H6').setValue(0.7732);
  ss.getRange('H7').setValue(0.056603987);
  ss.getRange('H8').setValue(0.3214);
  ss.getRange('H9').setValue(0.51925);
  ss.getRange('H10').setValue(0.0635 + 0.1574);
  ss.getRange('H11').setValue(0);
  ss.getRange('H12').setValue(1);
  ss.getRange('H13').setValue(1);
  ss.getRange('H14').setValue(1);
}

function onOpen(e) {
  //Initializes all things related to dates
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentDate = new Date();
  ss.getRange('I18').setValue(currentDate);
  
  //Stuff for extra menu item to Trace Dependents of a given cell
  var menuItems = [
    {name: 'Trace Dependents', functionName: 'traceDependents'}
  ];
//  menuEntries.push({name: "Trace Dependents", functionName: "traceDependents"});
//  menuEntries.push({name: "Trace Dependents", functionName: "traceDependents"});
  ss.addMenu("Detective", menuItems); 
}

function traceDependents(){
  var dependents = []
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentCell = ss.getActiveCell();
  var currentCellRef = currentCell.getA1Notation();
  var range = ss.getDataRange();

  var regex = new RegExp("\\b" + currentCellRef + "\\b");
  var formulas = range.getFormulas();

  for (var i = 0; i < formulas.length; i++){
    var row = formulas[i];

    for (var j = 0; j < row.length; j++){
      var cellFormula = row[j].replace(/\$/g, "");
      if (regex.test(cellFormula)){
        dependents.push([i,j]);
      }
    }
  }

  var dependentRefs = [];
  for (var k = 0; k < dependents.length; k ++){
    var rowNum = dependents[k][0] + 1;
    var colNum = dependents[k][1] + 1;
    var cell = range.getCell(rowNum, colNum);
    var cellRef = cell.getA1Notation();
    dependentRefs.push(cellRef);
  }
  var output = "Dependents: ";
  if(dependentRefs.length > 0){
    output += dependentRefs.join(", ");
  } else {
    output += " None";
  }
  currentCell.setNote(output);
}

