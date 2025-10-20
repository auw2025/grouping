// Import the necessary modules
const xlsx = require('xlsx');
const fs = require('fs');

// Read and parse the configuration from config.json (make sure the config file is in your project directory)
const config = JSON.parse(fs.readFileSync('student.config.json', 'utf8'));

// Destructure file paths from the configuration
const {
  userProfileFile,
  tableFile,
  classListFile,
  studentProfileFile
} = config.inputFiles;

const { studentProfileUpdated } = config.outputFiles;

// -------------------------------------------
// Process the user_profile_updated.xlsx file
// -------------------------------------------

// Read the user profile workbook (using config-specified file)
const userWorkbook = xlsx.readFile(userProfileFile);
// Assuming the data is in the first worksheet
const userSheetName = userWorkbook.SheetNames[0];
const userWorksheet = userWorkbook.Sheets[userSheetName];
// Convert the worksheet to JSON, using defval to capture empty cells as ""
const userJsonData = xlsx.utils.sheet_to_json(userWorksheet, { defval: "" });

// Build a mapping for all values in the user profile file.
// Instead of just a frequency map, store an object with a count and an array of sub_code values.
const frequencyMap = {};
userJsonData.forEach((row) => {
  // Iterate through each cell in the row
  Object.keys(row).forEach((columnName) => {
    // Get the cell value as a string and trim whitespace.
    const rawValue = row[columnName].toString().trim();
    if (rawValue) {
      // If the value contains a hyphen, remove it.
      const cellValue = rawValue.includes('-') ? rawValue.replace(/-/g, "") : rawValue;
      
      if (!frequencyMap[cellValue]) {
        frequencyMap[cellValue] = { count: 0, subCodes: [] };
      }
      frequencyMap[cellValue].count++;

      // If the row has a sub_code column and it's not empty, record it.
      if (row.hasOwnProperty('sub_code')) {
        const subCodeVal = row['sub_code'].toString().trim();
        if (subCodeVal) {
          frequencyMap[cellValue].subCodes.push(subCodeVal);
        }
      }
    }
  });
});

// Helper function to find matches in the user profile file for a given substring.
// It returns an object with the total count, matching keys information, and associated sub_code values.
const findMatches = (substring) => {
  let count = 0;
  let matches = [];
  const collectedSubCodes = new Set();

  // If the substring might contain a hyphen, remove it for matching purposes.
  const normalizedSubstring = substring.includes('-') ? substring.replace(/-/g, "") : substring;
  
  for (const key in frequencyMap) {
    // Check if the key contains the normalized substring.
    if (key.includes(normalizedSubstring)) {
      count += frequencyMap[key].count;
      matches.push(`${key} (${frequencyMap[key].count})`);
      frequencyMap[key].subCodes.forEach(sc => collectedSubCodes.add(sc));
    }
  }

  return { count, matches, subCodes: Array.from(collectedSubCodes) };
};

// -------------------------------------------
// Process the Table_Sub_Subject.xlsx file
// -------------------------------------------

// Read the Table_Sub_Subject workbook (using config-specified file)
const tableWorkbook = xlsx.readFile(tableFile);
// Assuming the data is in the first worksheet
const tableSheetName = tableWorkbook.SheetNames[0];
const tableWorksheet = tableWorkbook.Sheets[tableSheetName];
// Convert worksheet to JSON
const tableData = xlsx.utils.sheet_to_json(tableWorksheet, { defval: "" });

// Build a mapping from sub_code to subject_code
const subSubjectMapping = {};
tableData.forEach(row => {
  const subCode = row['sub_code'] ? row['sub_code'].toString().trim() : "";
  const subjectCode = row['subject_code'] ? row['subject_code'].toString().trim() : "";
  if (subCode && subjectCode) {
    subSubjectMapping[subCode] = subjectCode;
  }
});

// -------------------------------------------
// Helper Function: Convert sub codes if needed
// -------------------------------------------
const convertSubCode = (subCode) => {
  if (subCode === "DCHIS1") return "DCI11";
  if (subCode === "DCHIS3") return "DCI31";
  return subCode;
};

// -------------------------------------------
// Process the 2024-25 Class List.xlsx file
// -------------------------------------------

// Read the class list workbook (using config-specified file)
const classWorkbook = xlsx.readFile(classListFile);
// Assuming the data is in the first worksheet
const classSheetName = classWorkbook.SheetNames[0];
const classWorksheet = classWorkbook.Sheets[classSheetName];
// Convert the worksheet data to JSON format.
const classData = xlsx.utils.sheet_to_json(classWorksheet, { defval: "" });

// Array to hold new records that need to be appended to student_profile.xlsx
const appendedRecords = [];

// Iterate over each row in the class list data
classData.forEach((row, index) => {
  // Skip the row if columns X1, X2, and X3 are all empty
  if (!row['X1'] && !row['X2'] && !row['X3']) {
    return;
  }

  // Retrieve relevant values (defaulting to an empty string if missing)
  const tsssid = row['TSSSID'] || "";
  const form = row['Form'] || "";
  const x1 = row['X1'] || "";
  const x2 = row['X2'] || "";
  const x3 = row['X3'] || "";

  // Combine the Form with each X column to form new strings
  const formX1 = `${form}${x1}`;
  const formX2 = `${form}${x2}`;
  const formX3 = `${form}${x3}`;

  // Find matches for each of the combined values from the user_profile_updated.xlsx
  const matchX1 = findMatches(formX1);
  const matchX2 = findMatches(formX2);
  const matchX3 = findMatches(formX3);

  // Helper to build result string with sub_code and subject_code info for logging.
  const buildResultString = (matchObj) => {
    if (matchObj.count > 0) {
      // Convert each sub_code for display
      const convertedSubCodes = matchObj.subCodes.map(sc => convertSubCode(sc));
      const subCodesOutput = convertedSubCodes.length > 0 ? ` sub_code: ${convertedSubCodes.join(', ')}` : '';
      // Look up subject_code for the first converted sub_code found, if any
      const subjectCodesFound = convertedSubCodes
        .map(sc => subSubjectMapping[sc])
        .filter(sc => sc);  // filter out undefined or empty entries

      const subjectCodeOutput = subjectCodesFound.length > 0 ? ` subject_code: ${subjectCodesFound.join(', ')}` : '';
      return `Yes ${matchObj.count} (matched: ${matchObj.matches.join(', ')})${subCodesOutput}${subjectCodeOutput}`;
    } else {
      return "No";
    }
  };

  // Print the output for the current row
  console.log(`Row ${index + 1}:`);
  console.log(`  TSSSID: ${tsssid}`);
  console.log(`  FormX1: ${formX1} ${buildResultString(matchX1)}`);
  console.log(`  FormX2: ${formX2} ${buildResultString(matchX2)}`);
  console.log(`  FormX3: ${formX3} ${buildResultString(matchX3)}`);
  console.log('-------------------------');

  // For each Form field that resulted in Yes, create a new record.
  // If a row has more than one yes result, create one record per yes result.
  const processMatchForRecord = (matchObj, formValue) => {
    if (matchObj.count > 0 && matchObj.subCodes.length > 0 && matchObj.matches.length > 0) {
      // Use the first match to derive the grouping string.
      // For example, for "6CHIST2 BYC (3)" we take "6CHIST2 BYC"
      const grouping = matchObj.matches[0].split('(')[0].trim();
      // Use the first sub_code, and convert it if needed.
      const originalSub = matchObj.subCodes[0].toString().trim();
      const sub_code = convertSubCode(originalSub);
      // Look up the subject code with the modified sub_code value.
      const subject_code = subSubjectMapping[sub_code] || "";
      if (sub_code && subject_code) {
        // Create a new record object with 9 columns.
        return {
          tsss_id: tsssid,
          a_year: '2025/2026',
          terms: 'Term 1',
          sub_code: sub_code,
          subject_code: subject_code,
          taking: 'TRUE',
          self_taking: 'FALSE',
          grouping: grouping,
          grouping_real: grouping
        };
      }
    }
    return null;
  };

  // Process each form field
  const recordX1 = processMatchForRecord(matchX1, formX1);
  const recordX2 = processMatchForRecord(matchX2, formX2);
  const recordX3 = processMatchForRecord(matchX3, formX3);

  if (recordX1) {
    appendedRecords.push(recordX1);
  }
  if (recordX2) {
    appendedRecords.push(recordX2);
  }
  if (recordX3) {
    appendedRecords.push(recordX3);
  }
});

// -------------------------------------------
// Now, append the new records to the student_profile.xlsx data
// -------------------------------------------

// Read the original student_profile workbook (using config-specified file)
const studentWorkbook = xlsx.readFile(studentProfileFile);
const studentSheetName = studentWorkbook.SheetNames[0];
const studentWorksheet = studentWorkbook.Sheets[studentSheetName];
const studentData = xlsx.utils.sheet_to_json(studentWorksheet, { defval: "" });

// Append new records from appendedRecords
appendedRecords.forEach(record => {
  studentData.push(record);
});

// ---------------------------
// Sort the data by the "grouping" column in alphabetical order
// ---------------------------
studentData.sort((a, b) => {
  const groupingA = a.grouping ? a.grouping.toLowerCase() : "";
  const groupingB = b.grouping ? b.grouping.toLowerCase() : "";
  if (groupingA < groupingB) return -1;
  if (groupingA > groupingB) return 1;
  return 0;
});

// Create a new workbook that is a copy of the original student_profile.xlsx but with the appended records.
const newStudentWorkbook = xlsx.utils.book_new();
const newStudentWorksheet = xlsx.utils.json_to_sheet(studentData);
xlsx.utils.book_append_sheet(newStudentWorkbook, newStudentWorksheet, studentSheetName);

// Write the new workbook to a new file (using config-specified output file)
xlsx.writeFile(newStudentWorkbook, studentProfileUpdated);

console.log(`New student profile appended and sorted file created: '${studentProfileUpdated}'`);