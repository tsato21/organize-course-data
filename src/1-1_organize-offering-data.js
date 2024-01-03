const COURSES_NOT_CHECKED = ['MXD304','WNV327'];
/**
 * Organizes course offering data for display in a specified Google Sheet.
 * 
 * This function reads course offering data from one sheet and organizes it into a formatted structure on another sheet. 
 * It processes each row and column of the input data, filtering out specific course codes defined in `COURSES_NOT_CHECKED`.
 * The function skips empty cells and handles any errors during processing. The organized data includes the course code, 
 * offering semester, and faculty, starting from the third row of the destination sheet.
 * 
 * @param {string} programDataSheetName - Name of the Google Sheet containing the original course offering data.
 * @param {string} displaySheetName - Name of the Google Sheet where the organized data will be displayed.
 * 
 * Globals:
 * COURSES_NOT_CHECKED - Array of course codes that are not to be included in the processing.
 * 
 * Usage:
 * - Call `organizedGSCouseOfferingData` without any parameters.
 * - The function prompts for the names of the source and destination sheets.
 * - The processed data is displayed on the specified destination sheet.
 * 
 * Error Handling:
 * - Checks for empty sheet names and non-existent sheets.
 * - Catches and reports any runtime errors during processing.
 */
function organizedCourseOfferingDataType1() {
  let programDataSheetName = Browser.inputBox(`Input the name of the sheet where you pasted the course offering data provided by the Program (e.g., 1-1_Offering Data by Program).`,Browser.Buttons.OK_CANCEL);
  let programDataSheetNameTrimmed = programDataSheetName.trim();
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let programDataSheet = spreadSheet.getSheetByName(programDataSheetNameTrimmed);
  
  if (programDataSheetNameTrimmed === '') {
    Browser.msgBox("No sheet name was entered.");
    return;
  }
  if (programDataSheet === null) {
    Browser.msgBox("Sheet with the name '" + programDataSheetNameTrimmed + "' does not exist.");
    return;
  }

  let displaySheetName = Browser.inputBox(`Input the name of the sheet where organized course offering data is displayed (e.g., 1-1_Organized Offering Data).`,Browser.Buttons.OK_CANCEL);
  let displaySheetNameTrimmed = displaySheetName.trim();
  let displaySheet = spreadSheet.getSheetByName(displaySheetNameTrimmed);
  
  if (displaySheetNameTrimmed === '') {
    Browser.msgBox("No sheet name was entered.");
    return;
  }
  if (displaySheet === null) {
    Browser.msgBox("Sheet with the name '" + displaySheetNameTrimmed + "' does not exist.");
    return;
  }

  let data = programDataSheet.getDataRange().getValues();
  let newData = [];

  try{
      for (let i = 1; i < data.length; i++) { // Start from row 2 to skip the header row
      let faculty = data[i][0].toUpperCase();
      for (let j = 1; j<data[i].length; j++){
        if(data[i][j] === ''){
          continue;
        }
        let courseCode = data[i][j];
        let offeringSemester = data[0][j];
        console.log(`Course Code is ${courseCode} and semester is ${offeringSemester}`);

        if (courseCode && !COURSES_NOT_CHECKED.includes(courseCode)) {
          let newRow = [courseCode, offeringSemester, faculty];
          newData.push(newRow);
        }
      }
    }
    // console.log(newData);
    
    // Clear the original data and write the modified data back to the sheet
    let displaySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(displaySheetName);
    let columnAValues = displaySheet.getRange("A:A").getValues();
    let lastRow = columnAValues.filter(String).length;
    if (lastRow > 3) {
      displaySheet.getRange(3, 1, lastRow, 3).clearContent();
    }
    
    displaySheet.getRange(3,1, newData.length, 3).setValues(newData);
    Browser.msgBox("Offering data by the Program is displayed.");
  } catch(e){
    Browser.msgBox(`The folllowing error occured: ${e.message}`)
  }
  
}