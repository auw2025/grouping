/**
 * ================================================================
 *    Program to:
 *      1) Process "2425 Teacher Workload.xlsx" to create "user_profile.xlsx"
 *         - For the "MATHS" sheet, if column 7 data equals "C+M2", then produce an 
 *           additional concatenated row by replacing the second token with "(M2)"
 *           and using only the leading number from the first token.
 *      2) Print first two columns of "Table_Sub_Subject.xlsx"
 *      3) Read "user_profile.xlsx" - for each "grouping" value:
 *         - Derive a NEW subject code using your new logic
 *         - Check Table_Sub_Subject.xlsx's 2nd column for an exact or "contains" match
 *         - If a match is found, get the row's "sub_code" column value
 *         - Write that value back to the second column of user_profile.xlsx
 * ================================================================
 */

const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

// --- Load configuration ---
const configPath = path.join(__dirname, 'teacher.config.json');
// Using require() is simple and synchronous:
const config = require(configPath);
// Alternatively, you can load via fs if you need asynchronous or more complex logic.
// const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));

// Use file names from config.
const teacherWorkloadFile = config.teacherWorkloadFile;
const tableSubSubjectFile = config.tableSubSubjectFile;
const userProfileIntermediateFile = config.userProfileIntermediateFile;

// ===== PART 1: Process teacher workload file to generate user_profile.xlsx ===== //

const inputFilePath = path.join(__dirname, teacherWorkloadFile);
const workloadWorkbook = xlsx.readFile(inputFilePath);

// Helper: sheet name must be uppercase only, no underscores.
const isValidSheetName = (sheetName) => {
  if (sheetName.includes('_')) return false;
  return /^[A-Z]+$/.test(sheetName);
};

// Helper: ensure concatenated value has a capital letter, a digit, no '---', not empty, etc.
const isValidConcatenation = (str) => {
  const trimmed = str.trim();
  if (trimmed === '') return false;
  if (!/[A-Z]/.test(trimmed)) return false;
  if (!/\d/.test(trimmed)) return false;
  if (trimmed.includes('---')) return false;
  if (trimmed.includes('ACE')) return false;
  if (trimmed.includes('READING')) return false;
  return true;
};

const outputRows = [];
const subjectCodes = [];

// Process each eligible sheet in workloadWorkbook.
workloadWorkbook.SheetNames.forEach((sheetName) => {
  if (!isValidSheetName(sheetName)) return;

  console.log(`\nProcessing Sheet: ${sheetName}`);
  const sheet = workloadWorkbook.Sheets[sheetName];
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const headerRowIndex = 2; // Excel row 3 is index 2
  
  // Read header row (row 3).
  let headerValues = [];
  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellRef = xlsx.utils.encode_cell({ c: col, r: headerRowIndex });
    headerValues.push(sheet[cellRef] ? sheet[cellRef].v : '');
  }
  // Trim trailing empties.
  while (headerValues.length && headerValues[headerValues.length - 1] === '') {
    headerValues.pop();
  }
  console.log('Headers:', headerValues);

  console.log("Processing concatenated values of Column 4 & 5:");
  // Data rows start at row 4 => index 3.
  for (let row = headerRowIndex + 1; row <= range.e.r; row++) {
    const cellRefCol4 = xlsx.utils.encode_cell({ c: 3, r: row });
    const cellRefCol5 = xlsx.utils.encode_cell({ c: 4, r: row });

    const valCol4 = sheet[cellRefCol4] ? sheet[cellRefCol4].v : '';
    const valCol5 = sheet[cellRefCol5] ? sheet[cellRefCol5].v : '';

    // Replace multiple spaces with one space on each value before concatenation.
    const cleanValCol4 = String(valCol4).replace(/\s+/g, ' ');
    const cleanValCol5 = String(valCol5).replace(/\s+/g, ' ');

    const concatenated = `${cleanValCol4} ${cleanValCol5}`.trim();

    if (isValidConcatenation(concatenated)) {
      // Extract staff_code: substring after the first space, then if another space exists, take the part after it.
      let staff_code = '';
      const firstSpaceIdx = concatenated.indexOf(' ');
      if (firstSpaceIdx !== -1) {
        staff_code = concatenated.substring(firstSpaceIdx + 1).trim();
      }
      const innerSpaceIdx = staff_code.indexOf(' ');
      if (innerSpaceIdx !== -1) {
        staff_code = staff_code.substring(innerSpaceIdx + 1).trim();
      }

      // Original subject code extraction logic.
      const parts = concatenated.split(' ').filter((p) => p.trim() !== '');
      let subjectCode = null;
      if (parts.length === 3) {
        subjectCode = parts[1]; // use second token.
      } else if (parts.length === 2) {
        let candidate = parts[0];
        if (!/\d/.test(candidate)) {
          subjectCode = candidate;
        } else {
          subjectCode = candidate.replace(/\d/g, '');
        }
      }

      if (subjectCode) {
        subjectCodes.push(subjectCode);
        console.log(`Row ${row + 1}: Concatenated="${concatenated}" | Subject Code="${subjectCode}"`);
      } else {
        console.log(`Row ${row + 1}: Concatenated="${concatenated}" | Subject Code not derived`);
      }

      // Build output row (for the original concatenation).
      outputRows.push({
        user: staff_code,
        sub_code: '',         // Will be updated in Part 4.
        class: concatenated,
        grouping: concatenated,
        grouping_real: concatenated
      });

      // If this is the "MATHS" sheet, check column 7 for a "C+M2" value.
      if (sheetName === "MATHS") {
        const cellRefCol7 = xlsx.utils.encode_cell({ c: 6, r: row });
        const valCol7 = sheet[cellRefCol7] ? sheet[cellRefCol7].v : '';
        if (String(valCol7).trim().includes("C+M2")) {
          // Get the leading number from parts[0]. E.g. "4MJPWT5" becomes "4".
          let firstNumber = parts[0].match(/^\d+/);
          if (firstNumber) {
            firstNumber = firstNumber[0];
          } else {
            firstNumber = parts[0];
          }
          // Construct the extra concatenated string:
          // Use only the leading number from the first token.
          const extraConcatenated = `${firstNumber} (M2) ${parts.slice(2).join(" ")}`;
          console.log(`Row ${row + 1}: Concatenated="${extraConcatenated}" | Subject Code="M2"`);
          // Also add this extra row to be included in user_profile.xlsx.
          outputRows.push({
            user: staff_code,
            sub_code: '',  // This row's subject code will be set in Part 4.
            class: extraConcatenated,
            grouping: extraConcatenated,
            grouping_real: extraConcatenated
          });
        }
      }
    }
  }

  // Additionally, print any data in Column 7 (Index 6) for the "MATHS" sheet.
  if (sheetName === "MATHS") {
    console.log(`\nAdditional Check: Data in Column 7 (Index 6) in sheet "${sheetName}":`);
    for (let row = headerRowIndex + 1; row <= range.e.r; row++) {
      const cellRefCol7 = xlsx.utils.encode_cell({ c: 6, r: row });
      const valCol7 = sheet[cellRefCol7] ? sheet[cellRefCol7].v : '';
      if (String(valCol7).trim().includes('C+M2')) {
        console.log(`Row ${row + 1}: Column 7 Data = "${valCol7}"`);
      }
    }
  }
});

// Save user_profile.xlsx with the file name from config.
const outputWorkbook = xlsx.utils.book_new();
const outputSheet = xlsx.utils.json_to_sheet(outputRows, {
  header: ['user', 'sub_code', 'class', 'grouping', 'grouping_real']
});
xlsx.utils.book_append_sheet(outputWorkbook, outputSheet, 'UserProfile');

const outputFilePath = path.join(__dirname, userProfileIntermediateFile);
xlsx.writeFile(outputWorkbook, outputFilePath);
console.log(`\nUser profile output written to ${outputFilePath}`);

// Print all derived subject codes.
console.log("\nAll Derived Subject Codes:");
subjectCodes.forEach((subjectCode, idx) => {
  console.log(`${idx + 1}. ${subjectCode}`);
});

// ===== PART 2: Process Table_Sub_Subject.xlsx to print first two columns ===== //
const tableFilePath = path.join(__dirname, tableSubSubjectFile);
const tableWorkbook = xlsx.readFile(tableFilePath);
const tableSheetName = tableWorkbook.SheetNames[0];
const tableSheet = tableWorkbook.Sheets[tableSheetName];
const tableRange = xlsx.utils.decode_range(tableSheet['!ref']);

console.log(`\nProcessing Table_Sub_Subject.xlsx (Sheet: ${tableSheetName})`);

// Grab headings from columns 1 and 2.
const headingCol1Ref = xlsx.utils.encode_cell({ c: 0, r: 0 });
const headingCol2Ref = xlsx.utils.encode_cell({ c: 1, r: 0 });
const headingCol1 = tableSheet[headingCol1Ref] ? tableSheet[headingCol1Ref].v : '';
const headingCol2 = tableSheet[headingCol2Ref] ? tableSheet[headingCol2Ref].v : '';
console.log(`Headings: [ "${headingCol1}", "${headingCol2}" ]`);

// Print each row's col1 and col2.
for (let row = 1; row <= tableRange.e.r; row++) {
  const cellCol1Ref = xlsx.utils.encode_cell({ c: 0, r: row });
  const cellCol2Ref = xlsx.utils.encode_cell({ c: 1, r: row });

  const valCol1 = tableSheet[cellCol1Ref] ? tableSheet[cellCol1Ref].v : '';
  const valCol2 = tableSheet[cellCol2Ref] ? tableSheet[cellCol2Ref].v : '';

  console.log(`Row ${row + 1}: [ "${valCol1}", "${valCol2}" ]`);
}

// ===== PART 4: Read user_profile.xlsx and derive a NEW subject code from "grouping", then store to sub_code ===== //
const userProfilePath = path.join(__dirname, userProfileIntermediateFile);
const userProfileWorkbook = xlsx.readFile(userProfilePath);
const userProfileSheet = userProfileWorkbook.Sheets['UserProfile'];
let userProfileData = xlsx.utils.sheet_to_json(userProfileSheet);

console.log('\nPart 4: For each row in user_profile.xlsx, derive a NEW subject code and store the matched sub_code.');

// 4.1) Identify which column in Table_Sub_Subject.xlsx is named "sub_code" (in row 0 / header)
let subCodeColIndex = -1;  // Index of the "sub_code" column
{
  const headerRow = 0; // Typically the first row is the header.
  for (let c = 0; c <= tableRange.e.c; c++) {
    const cellRef = xlsx.utils.encode_cell({ c, r: headerRow });
    const val = tableSheet[cellRef] ? tableSheet[cellRef].v : '';
    if (typeof val === 'string' && val.trim().toLowerCase() === 'sub_code') {
      subCodeColIndex = c;
      break;
    }
  }
}

// Define the special mapping for DSE subjects.
const dseMapping = {
  "DSE-PE1": "DPED1",
  "DSE-PE2": "DPED2",
  "DSE-PE3": "DPED3",
  "DSE-VA1": "DVIAR1",
  "DSE-VA2": "DVIAR2",
  "DSE-VA3": "DVIAR3"
};

userProfileData.forEach((row, index) => {
  // index + 2 => Excel row numbering if the first row is header.
  const groupingVal = row['grouping'] || '';
  console.log(`\nRow ${index + 2}: grouping="${groupingVal}"`);

  // --- Apply the NEW subject-code extraction logic ---
  const parts = groupingVal.split(' ').filter((p) => p.trim() !== '');
  let newSubjectCode = '';

  if (parts.length === 3) {
    // e.g. "6MJPWT3 MATHS SUW" or "6 (M2) SUW" => use the second token.
    newSubjectCode = parts[1];
  } else if (parts.length === 2) {
    newSubjectCode = parts[0].replace(/^\d+/, '').trim();
    if (groupingVal.toUpperCase().includes('DSE')) {
      console.log(`  => 'DSE' detected. Derived subject code from first part: "${newSubjectCode}"`);
    } else {
      console.log(`  => 'DSE' not detected. Derived subject code from first part: "${newSubjectCode}"`);
    }
  }
  // Remove surrounding parentheses if present (e.g., "(M2)" becomes "M2").
  newSubjectCode = newSubjectCode.replace(/^\((.*)\)$/, '$1');

  // Apply special normalizations (for non-DSE cases).
  if (newSubjectCode === "MATHS") {
    newSubjectCode = "MATH";
  }
  if (newSubjectCode === "RSE") {
    newSubjectCode = "REGS";
  }
  if (newSubjectCode === "L&S") {
    newSubjectCode = "LISO";
  }
  if (newSubjectCode === "VA") {
    newSubjectCode = "VIAR";
  }
  if (newSubjectCode === "PTH") {
    newSubjectCode = "PUTO";
  }

  console.log(`  => NEW Subject Code = "${newSubjectCode}"`);

  // 4.2) Determine the matched sub_code.
  let foundMatch = false;
  let matchedSubCodeVal = '';

  // Special mapping for DSE subjects.
  if (groupingVal.toUpperCase().includes('DSE') && dseMapping[newSubjectCode]) {
    foundMatch = true;
    matchedSubCodeVal = dseMapping[newSubjectCode];
    console.log(`  => Special DSE mapping: For NEW Subject Code "${newSubjectCode}", using mapped sub_code: "${matchedSubCodeVal}"`);
  }
  // Special override for our M2 record.
  else if (newSubjectCode === "M2") {
    foundMatch = true;
    matchedSubCodeVal = "DMA21";
    console.log(`  => Special case for M2: Using "DMA21"`);
  } else if (newSubjectCode) {
    // Check for special CHIST* mappings.
    if (newSubjectCode.includes('CHIST1')) {
      foundMatch = true;
      matchedSubCodeVal = 'DCI11';
      console.log(`  => Special mapping: NEW Subject Code contains 'CHIST1', using "${matchedSubCodeVal}"`);
    } else if (newSubjectCode.includes('CHIST2')) {
      foundMatch = true;
      matchedSubCodeVal = 'DCI21';
      console.log(`  => Special mapping: NEW Subject Code contains 'CHIST2', using "${matchedSubCodeVal}"`);
    } else if (newSubjectCode.includes('CHIST3A')) {
      foundMatch = true;
      matchedSubCodeVal = 'DC3A1';
      console.log(`  => Special mapping: NEW Subject Code contains 'CHIST3A', using "${matchedSubCodeVal}"`);
    } else if (newSubjectCode.includes('CHIST3B')) {
      foundMatch = true;
      matchedSubCodeVal = 'DC3B1';
      console.log(`  => Special mapping: NEW Subject Code contains 'CHIST3B', using "${matchedSubCodeVal}"`);
    } else if (newSubjectCode.includes('CHIST3')) {
      foundMatch = true;
      matchedSubCodeVal = 'DCI31';
      console.log(`  => Special mapping: NEW Subject Code contains 'CHIST3', using "${matchedSubCodeVal}"`);
    } else if (newSubjectCode === "CHI") {
      foundMatch = true;
      matchedSubCodeVal = "CHN1";
      console.log(`  => Special case for CHI: Using first matched sub_code = "CHN1"`);
    } else if (newSubjectCode === "IS") {
      foundMatch = true;
      matchedSubCodeVal = "INSC";
      console.log(`  => Special case for IS: Using first matched sub_code = "INSC"`);
    } else {
      // Otherwise, search for an exact or "contains" match in the table.
      for (let rowT = 1; rowT <= tableRange.e.r; rowT++) {
        const cellRefB = xlsx.utils.encode_cell({ c: 1, r: rowT }); // Column B => subject_code.
        const bVal = tableSheet[cellRefB] ? tableSheet[cellRefB].v : '';

        if (bVal === newSubjectCode) {
          foundMatch = true;
          console.log(`  => EXACT match: Table_Sub_Subject row ${rowT + 1} => "${bVal}"`);
          if (subCodeColIndex > -1) {
            const cellRefSubCode = xlsx.utils.encode_cell({ c: subCodeColIndex, r: rowT });
            matchedSubCodeVal = tableSheet[cellRefSubCode] ? tableSheet[cellRefSubCode].v : '';
          }
          break;
        } else if (typeof bVal === 'string' && bVal.includes(newSubjectCode)) {
          foundMatch = true;
          console.log(`  => CONTAINS match: Table_Sub_Subject row ${rowT + 1} => "${bVal}"`);
          if (subCodeColIndex > -1) {
            const cellRefSubCode = xlsx.utils.encode_cell({ c: subCodeColIndex, r: rowT });
            matchedSubCodeVal = tableSheet[cellRefSubCode] ? tableSheet[cellRefSubCode].v : '';
          }
          break;
        }
      }
    }
  }

  if (!foundMatch && newSubjectCode) {
    console.log(`  => No match found in second column for "${newSubjectCode}".`);
  }

  console.log(`  => First matched sub_code = "${matchedSubCodeVal}"`);

  // 4.3) Store matchedSubCodeVal into the "sub_code" field of this row in userProfileData.
  row['sub_code'] = matchedSubCodeVal;  // If not found, remains "".
});

// 4.4) Sort the updated userProfileData by the "user" field before saving.
userProfileData.sort((a, b) => {
  const userA = String(a.user).toUpperCase();
  const userB = String(b.user).toUpperCase();
  return userA.localeCompare(userB);
});

// Overwrite user_profile.xlsx with the updated and sorted data.
const updatedSheet = xlsx.utils.json_to_sheet(userProfileData, {
  header: ['user', 'sub_code', 'class', 'grouping', 'grouping_real']
});
userProfileWorkbook.Sheets['UserProfile'] = updatedSheet;
xlsx.writeFile(userProfileWorkbook, userProfilePath);

console.log('\nDone! The updated "user_profile.xlsx" now has the matched sub_code values in column "sub_code", sorted by the "user" column.');