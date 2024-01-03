const FACULTY_NAME_REPLACEMENT = [
            {original: 'LUHFE', replace: 'LL'},
            ];
/**
 * Organizes class schedule data from a specified Google Sheets document.
 *
 * This function processes class schedule data from a provided sheet and organizes it into a structured format
 * on another sheet. It filters and formats the data, converting English day names to Japanese, 
 * replacing faculty names according to predefined rules, and handling multiple course codes and faculties.
 * The function also groups and sorts the data by course code and day before writing it to the target sheet.
 *
 * @param {string} programDataSheetName - Name of the Google Sheet containing the original class schedule data.
 * @param {string} displaySheetName - Name of the Google Sheet where the organized class schedule data will be displayed.
 * 
 * Usage:
 * - Call `organizedGBClassScheduleData` without any parameters.
 * - The function prompts for the names of the source and destination sheets.
 * - Organized data, including course code, course title, faculty, day, class number, start time, and end time,
 *   is displayed on the specified destination sheet.
 * 
 * Error Handling:
 * - Verifies the existence of the input and output sheets.
 * - Catches and reports runtime errors during data processing.
 */
function organizedClassScheduleDataType2() {
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let programDataSheetName = Browser.inputBox(`Input the name of the sheet where you pasted the class schedule data provided by the Program (e.g., 2_Schedule Data by Program).`,Browser.Buttons.OK_CANCEL);
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

  let displaySheetName = Browser.inputBox(`Input the name of the sheet where organized course offering data is displayed (e.g., 2_Organized Schedule Data).`,Browser.Buttons.OK_CANCEL);
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
  let outputData = [];
  
  // Updated header to include "Day"
  outputData.push(['Course Code', 'Course Title', 'Faculty','Day', 'Class Number', 'Start Time', 'End Time']);  
  
  let currentDay = '';  // This will hold the value of the current day as we iterate
  let intermediateData = [];

  // Dictionary to convert English day names to Japanese
  let dayConversion = {
      "Monday": "月",
      "Tuesday": "火",
      "Wednesday": "水",
      "Thursday": "木",
      "Friday": "金"
  };

  //Start time, end time, and each course details are described in the third row or after.
  //Loop to read data at each column
  for (let i = 2; i < data[0].length; i++) {
    // console.log(`i's length is ${data[0].length}`);
    if (data[0][i]) {
      // Update the current day by referring to the first row of the same column, converting to Japanese if applicable
      currentDay = dayConversion[data[0][i]] || data[0][i];
    }
    //Loop to read data at each cell vertically
    for (let j = 2; j < data.length; j++) {
        // console.log(`j's length is ${data.length}`);
        console.log(`data[j][i] is ${data[j][i]}`);
        // Check if the starting elements match the pattern (e.g. ECN345)
        let pattern = /^[A-Za-z]{3}\d{3}/;
        //if the cell has data and contains the designated pattern
        if (data[j][i] && pattern.test(data[j][i])) {
          //Refer to the the second column of the target row where class schedule is described (e.g. 09:00-10:15)
          let times = data[j][1].split('-');
          let startTime = times[0].trim();
          let endTime = times[1].trim();

          /*
          The cell has three lines, and its composition is as follows:
          "
          Course code (e.g. ECN300)
          Course title
          Faculty name
          "
          */
          let courseDetails = data[j][i].split('\n').map(function(detail) {
              //Trim whitespace characters from both the beginning and the end of a string.
              return detail.trim();
          });
          // console.log(courseDetails);
          /*
            Some courses might have w-code pattersn or multiple faculties. In that case, split
            e.g.) courseCodes => ECN200/ECN203
            e.g.) faculties => Faculty A, Faculty B
          */
          let courseCodes = courseDetails[0].split('/');
          let courseTitle = courseDetails[1];
          let faculties = replaceFacultyName_(courseDetails[2]);
          
          /*
            If either courseCodes include "/" or faculties contain ",", make multiple records
            e.g.) courseCodes => ECN200/ECN203, courseTitile=>Course A faculties => Faculty A, Faculty B
                  ↓
                  ・Record 1: ECN200, Course A, Faculty A
                  ・Record 2: ECN200, Course A, Faculty B
                  ・Record 3: ECN203, Course A, Faculty A
                  ・Record 4: ECN203, Course A, Faculty B
          */
          for (let k = 0; k < courseCodes.length; k++) {
              for (let l = 0; l < faculties.length; l++) {
                  intermediateData.push({
                      courseCode: courseCodes[k],
                      courseTitle: courseTitle,
                      faculty: faculties[l].trim(),
                      day: currentDay,
                      classNumber: data[j][0],
                      startTime: startTime,
                      endTime: endTime,
                  });
              }
          }
        }
    }
  }
  
  // Group by "day" and "faculty"
  let groupedData = {};
  // Loop through each item in the intermediateData array.
  // Example: If intermediateData = [{courseCode: "CS101", day: "Monday", faculty: "Smith"}, ...]
  for (let i = 0; i < intermediateData.length; i++) {
      // Construct a unique key for each combination of courseCode, day, and faculty.
      // This key will be used to group data in the groupedData object.
      let key = intermediateData[i].courseCode + intermediateData[i].day + "_" + intermediateData[i].faculty;
      // If this key doesn't exist in groupedData, initialize it with an empty array.
      // This ensures we have an array ready to push data into.
      if (!groupedData[key]) {
          groupedData[key] = [];
      }
      groupedData[key].push(intermediateData[i]);
  }
  // console.log(JSON.stringify(groupedData));
  
  // Process grouped data to merge records
  for (let key in groupedData) {
      let records = groupedData[key];
      // console.log(JSON.stringify(groupedData[key]));
      if (records.length > 1) {
      // Check if records should not be merged based on classNumber
          if (records.some(record => record.classNumber === 2) && records.some(record => record.classNumber === 3)) {
              // Handle the case when class numbers 2 and 3 should not be merged
              // Process each record individually
              records.forEach(record => {
                  outputData.push([record.courseCode, record.courseTitle, record.faculty, record.day, record.classNumber, record.startTime, record.endTime]);
              });
          } else {
              let earliestStartTime = records[0].startTime; //set the first class start time as earliestStartTime and then, compare the other class start time and replace if necessary
              let latestEndTime = records[0].endTime; //set the first class end time as lastestEndTime and then, compare the other class end time and replace if necessary
              
              for (let i = 1; i < records.length; i++) {
                  if (records[i].startTime < earliestStartTime) {
                      earliestStartTime = records[i].startTime;
                  }
                  if (records[i].endTime > latestEndTime) {
                      latestEndTime = records[i].endTime;
                  }
              }
              
              outputData.push([records[0].courseCode, records[0].courseTitle, records[0].faculty,records[0].day, records[0].classNumber, earliestStartTime, latestEndTime]);
          }
      } else {
          // Single record, no merging needed
          let record = records[0];
          outputData.push([record.courseCode, record.courseTitle, record.faculty, record.day, record.classNumber, record.startTime, record.endTime]);
      }
  }

  // console.log(JSON.stringify(outputData));

  // Extract the header row
  let headerRow = outputData[0];

  // Exclude the header row from sorting
  let dataRows = outputData.slice(1);

  // Sort dataRows by course code and then by day
  dataRows.sort(function(a, b) {
      // First compare by course code
      if (a[0] < b[0]) return -1;
      if (a[0] > b[0]) return 1;

      // If course codes are equal, then compare by day
      if (a[3] < b[3]) return -1;
      if (a[3] > b[3]) return 1;

      // If both course code and day are equal
      return 0;
  });

  // Reassemble the outputData with the sorted rows and the header
  outputData = [headerRow].concat(dataRows);

  // Clear contents in target sheet
  displaySheet.getRange("A2:G").clearContent();
  // Write the sorted data to the sheet
  displaySheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);

  Browser.msgBox("Course data by the Program is displayed (Sorted by Course Code & Day).");
}

/**
 * Replaces faculty names based on predefined replacement rules.
 * 
 * This helper function takes a faculty name as input and replaces it with another name as specified
 * in the FACULTY_NAME_REPLACEMENT array. It also handles cases where multiple faculty names are provided, 
 * splitting them by commas and applying the replacement to each name.
 *
 * @param {string} facultyName - The original faculty name(s), potentially including multiple names separated by commas.
 * @return {string[]} An array of faculty names after applying the replacement rules and converting them to uppercase.
 */
function replaceFacultyName_(facultyName) {
  let replacedName = String(facultyName);
  // console.log(replacedName);
  FACULTY_NAME_REPLACEMENT.forEach(replacement => {
    replacedName = replacedName.replace(new RegExp(replacement.original, 'g'), replacement.replace);
  });
  return replacedName.toUpperCase().split(',');
}
