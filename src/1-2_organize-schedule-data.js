/**
 * Organizes and displays class schedule data from a specified Google Sheets document.
 *
 * This function processes class schedule data from an input sheet and organizes it into a formatted structure on another sheet.
 * It reads the data row by row and column by column, checking each cell against a specific pattern (three alphabets followed by three digits).
 * If the cell data does not match the pattern, it is stored as an exceptional case. The function handles day conversion from English to Japanese
 * and groups data based on course code and day. It also sorts the organized data by course code and day before displaying it on the target sheet.
 *
 * @param {string} programDataSheetName - Name of the Google Sheet containing the original class schedule data.
 * @param {string} displaySheetName - Name of the Google Sheet where the organized class schedule data will be displayed.
 * 
 * Usage:
 * - Call `organizedGSClassScheduleData` without any parameters.
 * - The function prompts for the names of the source and destination sheets.
 * - Organized data (including course code, day, class number, start time, and end time) is displayed on the specified destination sheet,
 *   while exceptional cases are listed separately.
 * 
 * Error Handling:
 * - Verifies the existence of the input and output sheets.
 * - Catches and reports runtime errors during data processing.
 */
function organizedClassScheduleDataType1() {
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let programDataSheetName = Browser.inputBox(`Input the name of the sheet where you pasted the class schedule data provided by the Program (e.g., 1-2_Schedule Data by Program).`,Browser.Buttons.OK_CANCEL);
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

  let displaySheetName = Browser.inputBox(`Input the name of the sheet where organized course offering data is displayed (e.g., 1-2_Organized Schedule Data).`,Browser.Buttons.OK_CANCEL);
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
  let exceptionalCells = []; // Array to store information about exceptional cells

  // Updated header
  outputData.push(['Course Code', 'Day', 'Class Number', 'Start Time', 'End Time']);
  
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

  let pattern = /^[A-Za-z]{3}\d{3}$/; // Pattern for course code validation

  //Since the course details are described in C2 or after, the target data is data[1][2] or after
  //Loop to read data at each column
  for (let i = 2; i < data[0].length; i++) {
    if (data[0][i]) {
      // Update the current day by referring to the first row of the same column, converting to Japanese if applicable
      currentDay = dayConversion[data[0][i]] || data[0][i];
    }
    //Loop to read data at each cell vertically
    for (let j = 1; j < data.length; j++) {
      let cellData = data[j][i];
      console.log(`cellData is ${cellData}`);

      if (cellData && cellData.length === 6 && pattern.test(cellData)) {
        let times = data[j][1].split('-');
        let startTime = times[0].trim();
        let endTime = times[1].trim();
        
        intermediateData.push({
            courseCode: cellData,
            day: currentDay,
            classNumber: data[j][0],
            startTime: startTime,
            endTime: endTime,
        });

      } else {
        if(cellData){
          exceptionalCells.push([cellData.replace(/\n/g, "")]);
        }
      }
    }
  }
  
  // Group by "day"
  let groupedData = {};
  // Loop through each item in the intermediateData array.
  // Example: If intermediateData = [{courseCode: "CS101", day: "Monday"}, ...]
  for (let i = 0; i < intermediateData.length; i++) {
      // Construct a unique key for each combination of courseCode and day.
      // This key will be used to group data in the groupedData object.
      let key = intermediateData[i].courseCode + intermediateData[i].day;
      // If this key doesn't exist in groupedData, initialize it with an empty array.
      // This ensures we have an array ready to push data into.
      if (!groupedData[key]) {
          groupedData[key] = [];
      }
      groupedData[key].push(intermediateData[i]);
  }
  console.log(JSON.stringify(groupedData));
  
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
                  outputData.push([record.courseCode, record.day, record.classNumber, record.startTime, record.endTime]);
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
              
              outputData.push([records[0].courseCode, records[0].day, records[0].classNumber, earliestStartTime, latestEndTime]);
          }
      } else {
          // Single record, no merging needed
          let record = records[0];
          outputData.push([record.courseCode, record.day, record.classNumber, record.startTime, record.endTime]);
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
  displaySheet.getRange("A2:E").clearContent();
  displaySheet.getRange("G2:G").clearContent();
  // Write the sorted data to the sheet
  displaySheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);

  if (exceptionalCells.length > 0) {
    // Handle exceptional cells, e.g., log them or show a message
    displaySheet.getRange(3, 7, exceptionalCells.length, 1).setValues(exceptionalCells);
  }

  Browser.msgBox("Course data by the Program is displayed as follows: Usual Courses are in A~E column, being sorted by Course Code & Day; Exceptional Courses are in J column. ");
}