const sheet = 
  SpreadsheetApp
  .getActiveSpreadsheet()
  .getSheetByName('Sheet1');                            // define worksheet
const range = 
  sheet.getRange(2,1,sheet.getLastRow() - 1, 5);        // define data range
const data = 
  range.getValues();                                    // define data 

const b = 1;
const e = 4;
const colB = data.map(i => [i[b]]);                     // define colB data
const colE = data.map(i => [i[e]]);                     // define colE data
const rangeB = 
  sheet.getRange(2, b+1, sheet.getLastRow() -1, 1);     // define colB range
const rangeE = 
  sheet.getRange(2, e+1, sheet.getLastRow() - 1, 1);    // define colE range

// Testing purposes only
const validate = () => Logger.log(colB)

const onOpen = () => {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu("ACL Tools")
  .addItem("Delete Pink/Yellow","deleteColored")
  .addItem("Format Col B","replaceSymbols")
  .addItem("Format Col E","setColumnE")
  .addToUi();
};

// Send message to Spreadsheet UI
const alertMessage = (message="alert") => {
  SpreadsheetApp.getUi().alert(message);
}

// Delete Pink/Yellow
const deleteColored = () => {
  const toDelete = getColored();
  let count = 0;
  
  if (toDelete.length) for (let row of toDelete) {
    sheet.deleteRow(row);
    count++;
  }

  alertMessage(`${count} rows of ${toDelete.length} Deleted`);
}

// Format Column B
const replaceSymbols = () => {
  const pattern = /[\W\d]+$/gu;

  const newData = colB.map(row => {
    if (row[0].match(pattern)) {
      return [`${row[0].replace(pattern, '')}{*}`];
    }
    return [row[0]];
  })

  rangeB.setValues(newData);
  alertMessage('Starting symbols replaced');
}

// Format Column B
const setColumnE = () => {
  const newValues = colB.map(i => ['approved_state__sys']);
  rangeE.setValues(newValues);
  alertMessage('Column E set!');
}

// return a list of all rows where ColB is NOT white background
const getColored = () => {
  const coloredCells = [];
  const bgColors = range.getBackgrounds();

  bgColors.forEach((row, index) => {
    if (row[1] !== '#ffffff') {
      coloredCells.unshift(parseInt(index)+2);
    }
  })
  
  return coloredCells;
}

// LEGACY
// const findSymbols = () => {
//   // const pattern = /^[^‘'A-Za-z0-9~"]+/u;
//   const pattern = /[\W\d]+$/gu;
//   const symbolsRemoved = [];

//   for (let row of colB) {
//     if (row[0].match(pattern)) {
//       symbolsRemoved.push([`${row[0].replace(pattern, '')}{*}`]);
//       //Logger.log(row[0].replace(pattern,'')); 
//     }
//     if (!row[0].match(pattern)) {
//       symbolsRemoved.push([`${row[0]}{*}`]);
//     }
//   }

//   return symbolsRemoved;
// }
