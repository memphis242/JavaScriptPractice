//RIGHT NOW, I'M SETTING THIS UP TO RUN WHEN BUTTON IS CLICKED; NEED TO IMPROVE BY MAKING IT UPON EDIT OF A ROW
function infoCheckHighlight() {
  var startRow = 5;                            //STARTING ROW OF COLUMN 'O'
  var maxRow = 50;                             //CURRENT MAXIMUM ROW FOR COLUMN 'O'; WILL INCREASE IF LOG PASSES THAT ROW
  var column = 15;                             //THIS IS COLUMN 'O', WHICH HOLDS ALL THE CHECK VALUES
  var initialCol = 2;                          //THIS IS COLUMN 'B', WHICH IS THE COLUMN WHERE THE LOG STARTS
  var finalCol = column;                       //SAME AS COLUMN 'O', WHICH WOULD BE THE COLUMN WHERE THE LOG ENDS
  var dividerCol = 9;                          //THIS IS COLUMN 'I', WHICH DIVIDES THE TWO HALVES OF THE LOG
  var numOfCol = finalCol - initialCol;
  var green = '#00ff00';
  var red = '#dd7e6b';
  var darkRed = '#b31414';
  var blue = '#42edff'
  
  var currentSheets = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = currentSheets.getActiveSheet();
  var userInterface = SpreadsheetApp.getUi();  //FOR POPUPS
  var currentCell, currentCellVal, currentRow, rowCell;
  var missing;
  
  
  //GO THROUGH CELLS IN COLUMN 'O' TO CHECK
  for(i = startRow; i < maxRow; i++) {
    currentCell = sheet.getRange(i, column);
    currentCellVal = currentCell.getValue().toUpperCase();
    
    //IF YES, SET ROW TO GREEN BACKGROUND.
    if( currentCellVal == 'YES' ) {
      currentRow = sheet.getRange(i, initialCol, 1, numOfCol);
      currentRow.setBackground(green);
    }
    
    else if( currentCellVal == 'NO' ) {
      missing = true;
      
      //GO THROUGH CELLS IN ROW TO SEE WHERE INFORMATION IS MISSING
      for(j = initialCol; j <= finalCol; j++) {
        if( j == dividerCol ) {
          continue;
        }
        
        rowCell = sheet.getRange(i, j);
        
        //IF CELL IS EMPTY, SET RED TO INDICATE MISSING INFO.
        if( (rowCell.getValue() == '') || (rowCell.getValue() == '-') ) {
          rowCell.setBackground(darkRed);
        }
        //ELSE SET TO BLUE TO INDICATE THIS ROW HAS MISSING INFORMATION
        else {
          rowCell.setBackground(red);
        }
      }
    }
    
    //IF HALF OF THE INFO IS IN, WHICH INDICATES ITEM IS STILL IN LOST & FOUND
    else if( sheet.getRange(i, initialCol).getValue() != '' ) {
      currentRow = sheet.getRange(i, initialCol, 1, numOfCol);
      currentRow.setBackground(blue);
    }
    
    //ELSE JUST SET THE BACKGROUND OF THE ROW WHITE
    else {
      currentRow = sheet.getRange(i, initialCol, 1, numOfCol);
      currentRow.setBackground('white');
    }
    
  }
  
  //THIS IS INEFFICIENT BUT SINCE THE DIVIDIER COLUMN WAS PROBABLY COLORED, RESET IT TO BLACK
  sheet.getRange(startRow, dividerCol, maxRow, 1).setBackground('black');
  
  //IF THERE WAS ANY MISSING INFORMATION, POPUP
  if( missing ) {
    userInterface.alert('Please make sure to fill in all information. There is some missing.');
  }
  
}

