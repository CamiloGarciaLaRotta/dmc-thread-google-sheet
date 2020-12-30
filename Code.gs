// Allowlist of columns that contain DMC codes. Usually these will live in the main sheet.
// These will be the only columns for which a valid DMC code will result in appropriate colouring
const COLOUR_SHEET_NAME = 'colours';
const PINK_COL = 1;
const RED_COL = 3;
const ORANGE_COL = 5;
const YELLOW_COL = 7;
const GREEN_COL = 9;
const BLUE_COL = 11;
const PURPLE_COL = 13;
const BROWN_COL = 15;
const WHITE_COL = 17;
const VALID_COLS = [PINK_COL, RED_COL, ORANGE_COL, YELLOW_COL, GREEN_COL, BLUE_COL, PURPLE_COL, BROWN_COL, WHITE_COL];

// DMC code, DMC name and corresponding HEX colour columns.
// They could live in a separate sheet.
const CODE_SHEET_NAME = 'colour_codes';
const CODE_COL = 1;
const NAME_COL = 2;
const HEX_COL = 3;
var codeSheet = SpreadsheetApp.getActive().getSheetByName(CODE_SHEET_NAME)

// Non numerical DMC codes and their correspondent row in the CODE_SHEET_NAME
const SNOW_WHITE_ROW = 452;
const SNOW_WHITE_CODE = "b5200";

const ECRU_ROW = 453;
const ECRU_CODE = "ecru";

const WHITE_ROW = 454;
const WHITE_CODE = "white";

// If cell A1 is given a valid DMC code, colour A1 with the correspondent HEX code,
// and write in cell B1 the correspondent DMC name.
// It listens to every cell edit, only act on columns in the VALID_COLS allowlist
function onEdit(e) {
  Logger.log("column: " + e.range.getColumn());
  
  var colorSheet = e.source.getActiveSheet();
  if (colorSheet.getName() !== COLOUR_SHEET_NAME) return;
  
  if (inactionableEvent(e)) return;
  
  var inputValue = e.value.toString().toLowerCase();
  Logger.log("dmc code: " + inputValue);  
  
  var row = getRow(inputValue, codeSheet);
  Logger.log("row: " + row);  
  if (row === -1) return;
  
  var hex = "#" + getHex(row, codeSheet);
  Logger.log("hex: " + hex);
  
  var name = getName(row, codeSheet);
  Logger.log("name: " + name, codeSheet);
  
  e.range.setBackground(hex);
  
  // if you don't want to write the name of the DMC code, you can comment out these lines.
  var nameCell = colorSheet.getRange(e.range.getRow(), e.range.getColumn() + 1);
  nameCell.setValue(name);
  nameCell.setBackground(hex);
}

function getHex(row, sh) {
  return sh.getRange(row, HEX_COL).getValue();
}

function getName(row, sh) {
  return sh.getRange(row, NAME_COL).getValue();
}

// given a dmc code and a sheet, find the it's row in the sheet.
function getRow(dmc, sh) {
  var dataRange = sh.getDataRange();
  var values = dataRange.getValues();

  // get non numerical codes out of the way
  switch(dmc) {
    case SNOW_WHITE_CODE:
      return SNOW_WHITE_ROW;
    case ECRU_CODE:
      return ECRU_ROW;
    case WHITE_CODE:
      return WHITE_ROW;
  }
  
  if (isNaN(dmc)) return -1;
  
  // binary search
  var low = 1
  var high = values.length;
  var mid;
  var curr;
  while (low <= high) {
    mid = Math.floor(low + ((high - low) / 2));
    // Logger.log("low: " + low);
    // Logger.log("mid: " + mid);
    // Logger.log("high: " + high);
    
    curr = values[mid][0];
    // Logger.log("curr: " + curr);
    
    if (dmc == curr) {
      return mid + 1;
    }
    
    if (dmc > curr) {
      low = mid + 1
    } else {
      high = mid - 1
    }
  }
  
  return -1;
}


// only act on single-cell changes on columns to which we expect DMC codes to be added to.
function inactionableEvent(e) {
  return e.range.getWidth() !== 1 || e.range.getHeight() !== 1 || 
    !VALID_COLS.includes(e.range.getColumn()) || !e.value;
}
