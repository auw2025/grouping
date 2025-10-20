const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// --- Load configuration from config file ---
// Ensure your config file has a .json extension (e.g. teacher.config.json)
const configPath = path.join(__dirname, 'teacher.config.json');
// You can either use require if the file uses the .json extension:
const config = require(configPath);
// Or load using fs if preferred:
// const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));

// Use config values for files.
const intermediateFile = config.userProfileIntermediateFile;
const finalFile = config.userProfileFinalFile;

// Load the workbook using the intermediate file name from config.
const workbookPath = path.join(__dirname, intermediateFile);
const workbook = XLSX.readFile(workbookPath);

// Assuming you are working with the first sheet in the workbook
const firstSheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[firstSheetName];

// Convert the sheet to JSON. Setting defval to an empty string to ensure empty cells are considered.
const data = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

// Mapping of specific sub_code replacements for first-level invalid rows.
const firstLevelSubCodeMapping = {
  "DENG1": "ENN1",
  "DMA11": "MATH",
  "REGS": "DREST",
  "DPED11": "PHED",
  "DCLI21": "COLI",
  "CHN1": "DCHIN1",
  "DPHY11": "PHY1",
  "DVIA11": "VIAR",
  "DMUSI1": "MUSI",
  "DGEO11": "GEOG"
};

// Mapping of specific sub_code replacements for second-level verification.
const secondLevelSubCodeMapping = {
  "DMA11": "DMATH1",
  "DPED11": "DPHEDS",
  "DPED3": "DPED31",
  "DPED2": "DPED21",
  "DCHS3A": "DC3A1",
  "DCHS3B": "DC3B1",
  "DCHIS2": "DCI21",
  "DCHIN": "DCHIN1",
  "DVIAR2": "DVIA21",
  "DVIAR3": "DVIA31"
};

// Function to extract the class number from the class cell value.
// It matches one or more digits at the beginning of the string.
function extractClassNumber(classStr) {
  const match = String(classStr).trim().match(/^(\d+)/);
  return match ? parseInt(match[1], 10) : null;
}

// First-level verification function that validates a given row.
// It returns true if the row is invalid based on first-level checks.
function firstLevelVerification(row, index) {
  let isRowInvalid = false;

  // Retrieve the class value and sub_code value.
  // Assuming the column names are exactly "class" and "sub_code".
  const classCellValue = row['class'];
  let subCode = String(row['sub_code'] || "").trim();

  // First-level check: Is the class cell well formed?
  const classNumber = extractClassNumber(classCellValue);
  if (classNumber === null) {
    console.log(`FIRST-LEVEL | Row ${index + 2}: INVALID - Invalid or missing class number from value "${classCellValue}".`);
    isRowInvalid = true;
    // Even if the class is invalid, update sub_code if it matches the mapping.
    const replacement = firstLevelSubCodeMapping[subCode.toUpperCase()];
    if (replacement) {
      console.log(`FIRST-LEVEL | Row ${index + 2}: Updating sub_code from "${subCode}" to "${replacement}"`);
      row['sub_code'] = replacement;
    }
    return isRowInvalid;
  }

  // First-level check: Validate sub_code based on class number
  // Conditions:
  // 1. If class number <= 3, the sub_code must NOT start with the letter 'D'
  // 2. If class number > 3, the sub_code must start with the letter 'D'
  let isValid = false;
  if (classNumber <= 3) {
    isValid = !(subCode.startsWith('D') || subCode.startsWith('d'));
  } else {
    isValid = (subCode.startsWith('D') || subCode.startsWith('d'));
  }

  if (!isValid) {
    console.log(`FIRST-LEVEL | Row ${index + 2}: INVALID - (class number: ${classNumber}, sub_code: "${subCode}")`);
    isRowInvalid = true;

    // If the row is invalid, update the sub_code if a mapping exists.
    const replacement = firstLevelSubCodeMapping[subCode.toUpperCase()];
    if (replacement) {
      console.log(`FIRST-LEVEL | Row ${index + 2}: Updating sub_code from "${subCode}" to "${replacement}"`);
      row['sub_code'] = replacement;
    }
  }

  return isRowInvalid;
}

// Second-level verification function to check sub_code replacements.
function secondLevelVerification(row, index) {
  // Retrieve and trim the sub_code value.
  let subCode = String(row['sub_code'] || "").trim();
  const replacement = secondLevelSubCodeMapping[subCode.toUpperCase()];
  if (replacement) {
    console.log(`SECOND-LEVEL | Row ${index + 2}: Updating sub_code from "${subCode}" to "${replacement}"`);
    row['sub_code'] = replacement;
  }
}

// Perform first-level verification on all rows.
let firstLevelInvalidCount = 0;
data.forEach((row, index) => {
  if (firstLevelVerification(row, index)) {
    firstLevelInvalidCount++;
  }
});
console.log(`\nTotal number of rows with FIRST-LEVEL invalidations: ${firstLevelInvalidCount}`);

// Now perform second-level verification on the rows.
data.forEach((row, index) => {
  secondLevelVerification(row, index);
});

// Convert the updated JSON data back to a worksheet.
const newWorksheet = XLSX.utils.json_to_sheet(data);

// Create a new workbook and append the new worksheet.
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, firstSheetName);

// Save the updated workbook to a new file using the final file name from config.
const newWorkbookPath = path.join(__dirname, finalFile);
XLSX.writeFile(newWorkbook, newWorkbookPath);

console.log(`Updated workbook has been saved as "${finalFile}".`);