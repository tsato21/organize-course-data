//Constant variables modifiable for your institution
const EAA_COURSES = ['EAA072', 'EAA082', 'EAA074', 'EAA084', 'EAA073', 'EAA083', 'EAA079', 'EAA089', 'EAA104','EAA105','EAA106','EAA096'];
const HIGHLIGHT_COURSES = ['EAA072', 'EAA082', 'EAA074', 'EAA084','EAA073', 'EAA083', 'EAA079', 'EAA089', 'EAA096','CCC100',];
const FACULTY_NAME_REPLACE = [
                    {preName: 'H.R', aftName: 'Hr RR'},
                    {preName: 'G.F', aftName: 'Gg FFFF'},
                    ];
const SECOND_RECORD_REMOVE_COURSES = ['CCC100'];

//Fixed constant variables
const MON_WED_COL_START = 2;
const MON_WED_COL_END = 23;
const TUE_THU_COL_START = 25;
const TUE_THU_COL_END = 46;
const FRI_COL_START = 48;
const FRI_COL_END = 61;

/**
 * Organizes class schedule data.
 * Prompts the user to input the names of the source data sheet and the display sheet.
 * Processes the class schedule data and outputs organized data to the display sheet.
 */
function organizedClassScheduleDataType3() {
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let programDataSheetName = Browser.inputBox(`Input the name of the sheet where you pasted the class schedule data provided by the Program (e.g., 2-3_Schedule Data by Program).`,Browser.Buttons.OK_CANCEL);
  let programDataSheetNameTrimmed = programDataSheetName.trim();
  let programDataSheet = spreadSheet.getSheetByName(programDataSheetNameTrimmed);
  
  if (programDataSheetNameTrimmed === '') {
    Browser.msgBox("No sheet name was entered.");
    return;
  }
  if (programDataSheet === null) {
    Browser.msgBox("Sheet with the name '" + programDataSheetNameTrimmed + "' does not exist.");
    return;
  }

  let displaySheetName = Browser.inputBox(`Input the name of the sheet where organized course offering data is displayed (e.g., 2-3_Organized Schedule Data).`,Browser.Buttons.OK_CANCEL);
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
  // let programDataSheet = spreadSheet.getSheetByName("Program CL Data Sheet"); // Replace with actual name
  // let displaySheet = spreadSheet.getSheetByName("Organized Program CL Data Sheet"); // Replace with actual name
  let data = programDataSheet.getDataRange().getValues();
  let semester = Browser.inputBox(`【★】Input Semester (Either 'S' for Spring Semester or 'F' for Fall Semester)`);
  
  if (semester === '') {
    Browser.msgBox(`The semester code is not input. Try again.`);
    return;
  } else if (semester !== 'S' && semester !== 'F'){
    Browser.msgBox(`The semester code should be either 'S' or 'F'. Try again.`);
    return;
  }

  let semesterCode = '_' + semester;
  let records = [];
  let header = ['Raw Value', 'Class Code', 'Faculty', 'Days of Week', 'Start Time', 'End Time', 'Room'];
  let sectionCodes = data[2].slice(1).map(extractAlphabets_); // Extract section codes from the third row

  for (let i = 3; i < data.length; i++) { // Skip the first two rows
    let row = data[i];
    for (let j = 1; j < row.length; j++) { // Start from the second column
      if (row[j] === '') continue;
      let rawValue = row[j]; 
      let modifiedRawValue = row[j].replace(/ /g, ""); // Use the replace method to modify the string in 'row[j]'.
      let cellColumn = j + 1; // +1 as column index starts from 1 in Sheets
      let lines = modifiedRawValue.split("\n").filter((line, index) => {
        // Exclude the line if it's the first line and empty, or if it contains the word "day" (case insensitive)
        return !(index === 0 && line.trim() === '') && !line.toLowerCase().includes("day");
      }); 

      let courseCode = lines[0] ? lines[0] : "";
      let facultyRawValue = lines[3] ? lines[3] : ""
      let faculty = facultyNameReplacement_(facultyRawValue);
      let timeMatch = modifiedRawValue.match(/(\d+:\d+)\s*-\s*(\d+:\d+)/);
      let startTime = timeMatch ? convertTimeFormat_(timeMatch[1]) : "";
      let endTime = timeMatch ? convertTimeFormat_(timeMatch[2]) : "";
      let room = (lines.length === 6 && lines[5] !== "") ? lines[5] : (lines[4] !== "" ? lines[4] : "");
      
      console.log(`Target record is as follows: courseCode is ${courseCode}; cellColumn is ${cellColumn}, sectionCode is ${sectionCodes[j - 1]}, faculty is ${faculty}, startTime is ${startTime}, endTime is ${endTime}, room is ${room}`);
      
      processCourseRecord_(courseCode, cellColumn, rawValue, sectionCodes[j - 1], semesterCode, faculty, startTime, endTime, room, records);
    }
  }
  displayRecords_(records, header, displaySheet);
}

/**
 * Extracts alphabet characters from a given string.
 * If the string includes a slash, only the first alphabet character before the slash is returned.
 * @param {string} inputString - The string to extract alphabets from.
 * @returns {string} The extracted alphabet characters.
 */
function extractAlphabets_(inputString) {
    // If the string includes "/", extract only the first alphabet character before "/"
    if (inputString.includes("/")) {
        return inputString.match(/[a-zA-Z]/)[0];
    }
    // Otherwise, match all alphabet characters and join them
    return (inputString.match(/[a-zA-Z]/g) || []).join('');
}

/**
 * Converts time from a string format to a standardized time format.
 * Adds a leading zero to hours if necessary.
 * @param {string} timeString - The time string to convert.
 * @returns {string} The converted time in HH:MM format.
 */
function convertTimeFormat_(timeString) {
  // Split the time string by colon
  let parts = timeString.split(":");

  if (parts[0] === "9") {
    parts[0] = "0" + parts[0];
  }

  // Return the hour and minute parts joined by a colon
  return parts[0] + ":" + parts[1];
}

/**
 * Processes each course record and categorizes it based on specific criteria.
 * Handles different course types including ADV course, EAA, and regular courses.
 * @param {string} courseCode - The course code.
 * @param {number} cellColumn - The column index of the cell.
 * @param {string} rawValue - The raw value from the cell.
 * @param {string} sectionCode - The section code for the course.
 * @param {string} semesterCode - The semester code (e.g., '_S' or '_F').
 * @param {string} faculty - The name of the faculty.
 * @param {string} startTime - The start time of the course.
 * @param {string} endTime - The end time of the course.
 * @param {string} room - The room number.
 * @param {Array} records - The array to store processed records.
 */
function processCourseRecord_(courseCode, cellColumn, rawValue, sectionCode, semesterCode, faculty, startTime, endTime, room, records) {
  let classCode;
  // Check if courseCode includes any string from designated types
  let slashIncludeCourse = courseCode.includes("/");
  let eaaCouse = EAA_COURSES.some(course => courseCode.includes(course));

  if (slashIncludeCourse) { //If the courseCode includes "/", split by "/" and push two records with some adjustment
    let parts = courseCode.split('/');
    // Check if there are two parts after splitting
    let courseCode_1 = parts[0];
    let courseCode_2 = "EAA" + parts[1];
    
    let classCode_1 = courseCode_1 + '-' + sectionCode + semesterCode;
    let classCode_2 = courseCode_2 + '-' + sectionCode + semesterCode;

    console.log(`${slashIncludeCourse} is splitted by "/" and courseCode is ${classCode_1} and ${classCode_2}.`)

    //Push the first record
    pushCourseRecord_(records, cellColumn, rawValue, classCode_1, faculty, startTime, endTime, room);
    //Push the second record
    pushCourseRecord_(records, cellColumn, rawValue, classCode_2, faculty, startTime, endTime, room);
    
  } else if (eaaCouse) { //If the courseCode is included in "EAA_COURSES" array, push the record with some adjustment
    classCode = courseCode + '-' + sectionCode + semesterCode;

    console.log(`${courseCode} is an EAA course.`);

    pushCourseRecord_(records, cellColumn, rawValue, classCode, faculty, startTime, endTime, room);

  } else { // If the courseCode is a regular one, push the record with some adjustment
    let match = courseCode.match(/-\d/);
    if(!match){
      classCode = courseCode + "-1" + semesterCode;
    } else {
      classCode = courseCode + semesterCode;
    }
    console.log(`${courseCode} is a regular course.`);
    pushCourseRecord_(records, cellColumn, rawValue, classCode, faculty, startTime, endTime, room);
  }
}

/**
 * Pushes a course record into the records array.
 * Determines the days of the week based on the column number and handles exceptions.
 * @param {Array} records - The array to store processed records.
 * @param {number} cellColumn - The column index of the cell.
 * @param {string} rawValue - The raw value from the cell.
 * @param {string} classCode - The class code for the course.
 * @param {string} faculty - The name of the faculty.
 * @param {string} startTime - The start time of the course.
 * @param {string} endTime - The end time of the course.
 * @param {string} room - The room number.
 */
function pushCourseRecord_(records, cellColumn, rawValue, classCode, faculty, startTime, endTime, room) {
  // Define the days based on the column range
  let daysOfWeek = [];
  let secondRecordRemoveCourse = SECOND_RECORD_REMOVE_COURSES.some(course => classCode.includes(course));
  if (cellColumn >= MON_WED_COL_START && cellColumn <= MON_WED_COL_END) {
      daysOfWeek = ['月','水'];
      console.log(`Since this cell(${cellColumn}) is in Monday/Wednesday range (${MON_WED_COL_START}:${MON_WED_COL_END}), record for Monday is pushed.`);
      if (secondRecordRemoveCourse) {
        daysOfWeek = ['月'];
        console.log(`Record for Wednesday is NOT pushed for ${classCode}.`);
      }
    } else if (cellColumn >= TUE_THU_COL_START && cellColumn <= TUE_THU_COL_END) {
      daysOfWeek = ['火','木'];
      console.log(`Since this cell(${cellColumn}) is in Tuesday/Thursday range(${TUE_THU_COL_START}:${TUE_THU_COL_END}), record for Tuesday is pushed.`);
      if (secondRecordRemoveCourse) {
        daysOfWeek = ['木'];
        console.log(`Record for Thursday is NOT pushed for ${classCode}.`);
      }
    } else if (cellColumn >= FRI_COL_START && cellColumn <= FRI_COL_END) {
      daysOfWeek = ['金'];
      console.log(`Since this cell(${cellColumn}) is in Friday range (${FRI_COL_START}:${FRI_COL_END}), record for Friday is pushed.`);
    } else {
      daysOfWeek = ['範囲外'];
      console.log(`The cell column of the record, ${rawValue} is ${cellColumn}, which is not within any of the designated ranges. Record as exception is pushed`);
    }

  // Push records for each day
  daysOfWeek.forEach(day => {
    records.push([rawValue, classCode, faculty, day, startTime, endTime, room]);
  });
}

/**
 * Displays the processed course records in the designated Google Sheet.
 * Clears existing content, formats the sheet, and highlights cells based on validation checks.
 * @param {Array} records - The array of processed course records.
 * @param {Array} header - The header row for the display.
 * @param {Object} displaySheet - The Google Sheet object where records will be displayed.
 */
function displayRecords_(records,header,displaySheet){
  displaySheet.getRange("A2:G").clearContent();

  // Get the range A2:G in the displaySheet
  let range = displaySheet.getRange("A3:G");
  // Clear font weight and background color in the specified range
  range.setFontWeight(null);  // Clears font weight
  range.setBackground(null);  // Clears background color

  displaySheet.getRange("A3:G").setBackground('#E7FCDB');
  displaySheet.getRange(2,1,1,header.length)
                                      .setValues([header]);

  // Write records to the sheet starting from row 2
  displaySheet.getRange(3, 1, records.length, header.length).setValues(records);
  //Set number format for start time and end time
  displaySheet.getRange(3,5,records.length,2).setNumberFormat("@");

  //Structure of each column
  let courseCodePattern = /^[A-Z]{3}\d{3}-[0-9A-Za-z]_[FS]$/;
  let facultyPattern = /^[A-Za-z]+$/; //Faculty column
  let dayPattern = /^(月|火|水|木|金)$/; //Days of week column
  let timePattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$/; //Start and end time column
  let roomPattern = /^[A-Z][0-9]{3}$/; //Room column
  
  // Iterate through the records to check for cells with "/" or empty cells (irregular cells)
  for (let i = 0; i < records.length; i++) {
    if (!courseCodePattern.test(records[i][1]) || records[i][1] == '') {
      highlightCells_(i + 2, 2,records[i][1], 'not valid for course code',displaySheet);
    } else if (HIGHLIGHT_COURSES.some(course => records[i][1].includes(course))) {
      highlightCells_(i + 2, 2,records[i][1], 'a course that needs to be checked',displaySheet);
    }
    if (!facultyPattern.test(records[i][2]) || records[i][2] == '') {
      highlightCells_(i + 2, 3,records[i][2], 'not valid for faculty',displaySheet);
    }
    if (!dayPattern.test(records[i][3]) || records[i][3] == '') {
      highlightCells_(i + 2, 4,records[i][3], 'not valid for days of week',displaySheet);
    }
    if (!timePattern.test(records[i][4]) || !timePattern.test(records[i][5]) || records[i][4] == '' || records[i][5] == '') {
        if (!timePattern.test(records[i][4])) {
          highlightCells_(i + 2, 5,records[i][4], 'not valid for start time',displaySheet);
        }
        if (!timePattern.test(records[i][5])) {
          highlightCells_(i + 2, 6,records[i][5], 'not valid for end time',displaySheet);
        }
    }
    if (!roomPattern.test(records[i][6]) || records[i][6] == '') {
      highlightCells_(i + 2, 7,records[i][6], 'not valid for room',displaySheet);
    }
  }
  let sheetName = displaySheet.getName();
  facultyNameReplacement_(displaySheet);

  Browser.msgBox(`Course data is output in ${sheetName} in the designated structure along with designated faculty names replaced.`);
}

/**
 * Highlights cells in the display sheet based on validation checks.
 * Changes the background color and font weight for cells with invalid data.
 * @param {number} rowNum - The row number of the cell to highlight.
 * @param {number} colNum - The column number of the cell to highlight.
 * @param {string} value - The value in the cell.
 * @param {string} highlightReason - The reason for highlighting the cell.
 * @param {Object} displaySheet - The Google Sheet object where the cell is located.
 */
function highlightCells_(rowNum, colNum, value, highlightReason,displaySheet){
  displaySheet.getRange(rowNum, colNum).setBackground("red").setFontWeight('bold');
  console.log(`The cell is colored in red since ${value} is ${highlightReason} or just empty.`);
}

/**
 * Replaces faculty names based on a predefined list of replacements.
 * @param {string} faculty - The original faculty name.
 * @returns {string} The replaced faculty name or the original name if no replacement is found.
 */
function facultyNameReplacement_(faculty) {  
  for (let j = 0; j < FACULTY_NAME_REPLACE.length; j++) {
    if (faculty === FACULTY_NAME_REPLACE[j].preName) {
        faculty = FACULTY_NAME_REPLACE[j].aftName;
        console.log(`Faculty name replaced: ${faculty}`);
        break; // Break the loop once a replacement is made
    }
  }
  return faculty;
}