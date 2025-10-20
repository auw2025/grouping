const XLSX = require('xlsx');
const config = require('./student.config.json');

// Function to load an Excel file and return its first sheet as JSON
function loadExcelFile(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(worksheet, { defval: "" });
}

// Helper function: Check if the row’s "class" is eligible.
// Eligible if exactly 3 words and first token’s numeric part is between 1 and 6.
function isEligibleClass(classStr) {
  if (typeof classStr !== "string") return false;
  const words = classStr.trim().split(/\s+/);
  if (words.length !== 3) return false;
  const firstComponent = words[0];
  const numericMatch = firstComponent.match(/^\d+/);
  if (!numericMatch) return false;
  const numericValue = parseInt(numericMatch[0], 10);
  return numericValue >= 1 && numericValue <= 6;
}

// Helper function: Extract form/class token (e.g. "1J" from "1J CES JAT").
function extractFormClass(classStr) {
  if (typeof classStr !== "string") return "";
  const tokens = classStr.trim().split(/\s+/);
  return tokens[0] || "";
}

/**
 * Search function used with the "2024-25 Class List.xlsx" data.
 * It will check the stored missing information (missingInfoMap) and
 * try to find an entry matching the following criteria:
 *
 *   - The stored missing info's group starts with the same number as the provided form.
 *   - The stored missing info's group includes the provided class letter.
 *   - The stored missing info's subject exactly matches the provided subject parameter.
 *   - The stored missing info's teacher exactly matches the provided teacher parameter.
 *
 * Returns the first matching info object (if found).
 *
 * @param {Map} missingInfoMap - Map with key=group and value = array of { group, subject, teacher, sub_code }.
 * @param {string|number} form - The Form value from the class list row.
 * @param {string} cls - The Class value (a letter) from the class list row.
 * @param {string} subject - The subject name to match.
 * @param {string} teacher - The teacher name to match.
 * @returns {Object|null} - The matching info object if found; otherwise, null.
 */
function searchMissingInfo(missingInfoMap, form, cls, subject, teacher) {
  const formStr = String(form);
  // iterate over each key-value pair; key is group token, value is an array of info objects
  for (const [group, entries] of missingInfoMap) {
    if (group.charAt(0) === formStr && group.includes(cls)) {
      // For each entry in the array, check if the subject and teacher match
      for (const info of entries) {
        if (info.subject === subject && info.teacher === teacher) {
          return info;
        }
      }
    }
  }
  return null;
}

function main() {
  // Retrieve file paths from configuration
  const { userProfileFile, tableFile, classListFile } = config.inputFiles;
  const { studentProfile } = config.outputFiles;

  // Load data from Excel files using files defined in the config
  const userData = loadExcelFile(userProfileFile);
  const tableData = loadExcelFile(tableFile);
  const classListData = loadExcelFile(classListFile);

  // Build a lookup map from sub_code to subject_code from Table_Sub_Subject.xlsx
  const subCodeMap = {};
  tableData.forEach(row => {
    // Assuming column names "sub_code" and "subject_code"
    subCodeMap[row.sub_code] = row.subject_code;
  });

  // Prepare arrays for output data and unprocessed row logs.
  const outputRows = [];
  const unprocessedRows = [];

  // Use a Map to store the missing class tokens information.
  // Key = group (first token), Value = array of { group, subject, teacher, sub_code }
  const missingInfoMap = new Map();

  // Process each row from user_profile_updated.xlsx
  userData.forEach((row, index) => {
    // Compute the Excel row number (assuming header is in row 1)
    const rowNumber = index + 2; 

    // Check eligibility of the 'class' field.
    if (!isEligibleClass(row.class)) {
      unprocessedRows.push({
        rowNumber,
        reason: `Ineligible class format. Expected exactly 3 tokens with the first token's numeric part between 1 and 6. Value found: '${row.class}'`
      });
      return; // Skip further processing for this row.
    }

    // Row is eligible: extract the form/class token.
    const formClassToken = extractFormClass(row.class);

    // Check if the form/class token exists in the class list.
    // Combine the Form and Class from the class list (trimmed) to form a token.
    const matchingClassRows = classListData.filter(item => {
      const combined = String(item.Form).trim() + String(item.Class).trim();
      return combined === formClassToken;
    });

    // If there is no matching class, store additional tokens.
    if (matchingClassRows.length === 0) {
      const tokens = row.class.trim().split(/\s+/);
      const group = tokens[0] || "";
      const subject = tokens[1] || "";
      const teacher = tokens[2] || "";
      // Also store the sub_code from the user data for later lookup.
      const sub_code = row.sub_code;

      // Instead of a single object per group, store an array of entries.
      if (!missingInfoMap.has(group)) {
        missingInfoMap.set(group, []);
      }
      // Push the new missing info even if one already exists for this group.
      missingInfoMap.get(group).push({ group, subject, teacher, sub_code });

      unprocessedRows.push({
        rowNumber,
        reason: `Eligible row but no matching class found in class list for token: '${formClassToken}'. Subject: '${subject}', Teacher: '${teacher}'`
      });
      return; // Skip further output row creation.
    }

    // Row is eligible and there is a matching class list record.
    // Retrieve subject_code using the sub_code lookup.
    const subject_code = Object.hasOwnProperty.call(subCodeMap, row.sub_code) ? subCodeMap[row.sub_code] : null;

    // Sort the matching rows by "Class no" numerically.
    matchingClassRows.sort((a, b) => {
      const numA = parseInt(a["Class no"], 10) || 0;
      const numB = parseInt(b["Class no"], 10) || 0;
      return numA - numB;
    });

    // For each matching class list row, push a new output row.
    matchingClassRows.forEach(matchRow => {
      outputRows.push({
        tsss_id: matchRow.TSSSID,
        a_year: "2025/2026",
        terms: "Term 1",
        sub_code: row.sub_code,
        subject_code: subject_code,
        taking: "TRUE",
        self_taking: "FALSE",
        grouping: row.class,
        grouping_real: row.class
      });
    });
  });

  // First, search stored missing info for each Class List row for subject "CHI"
  console.log("\nSearching stored missing info for each Class List row for subject CHI:");
  classListData.forEach((row, index) => {
    const missingEntry = searchMissingInfo(missingInfoMap, row.Form, row.Class, "CHI", row.CHI);
    if (missingEntry) {
      console.log(`Row ${index + 1} -> Matching group found for CHI: ${missingEntry.group}`);

      // Create a combined string "group subject teacher" e.g. "3MJPW4 CHI NEW"
      const combinedGrouping = `${missingEntry.group} ${missingEntry.subject} ${missingEntry.teacher}`;
      const subject_code = Object.hasOwnProperty.call(subCodeMap, missingEntry.sub_code) ? subCodeMap[missingEntry.sub_code] : null;

      // Add a new output record using this matching info.
      outputRows.push({
        tsss_id: row.TSSSID,
        a_year: "2025/2026",
        terms: "Term 1",
        sub_code: missingEntry.sub_code,
        subject_code: subject_code,
        taking: "TRUE",
        self_taking: "FALSE",
        grouping: combinedGrouping,
        grouping_real: combinedGrouping
      });
    } else {
      console.log(`Row ${index + 1} -> No matching group found for CHI`);
    }
  });

  // Next, search stored missing info for each Class List row for subject "MATHS"
  // The teacher field is taken from the "MATHS" column in the class list.
  console.log("\nSearching stored missing info for each Class List row for subject MATHS:");
  classListData.forEach((row, index) => {
    const missingEntry = searchMissingInfo(missingInfoMap, row.Form, row.Class, "MATHS", row.MATHS);
    if (missingEntry) {
      console.log(`Row ${index + 1} -> Matching group found for MATHS: ${missingEntry.group}`);

      // Create a combined string "group subject teacher" e.g. "3MJPW4 MATHS NEW"
      const combinedGrouping = `${missingEntry.group} ${missingEntry.subject} ${missingEntry.teacher}`;
      const subject_code = Object.hasOwnProperty.call(subCodeMap, missingEntry.sub_code) ? subCodeMap[missingEntry.sub_code] : null;

      // Add a new output record using this matching info.
      outputRows.push({
        tsss_id: row.TSSSID,
        a_year: "2025/2026",
        terms: "Term 1",
        sub_code: missingEntry.sub_code,
        subject_code: subject_code,
        taking: "TRUE",
        self_taking: "FALSE",
        grouping: combinedGrouping,
        grouping_real: combinedGrouping
      });
    } else {
      console.log(`Row ${index + 1} -> No matching group found for MATHS`);
    }
  });

  // Next, search stored missing info for each Class List row for subject "RSE"
  // The teacher field is taken from the "RSE" column in the class list.
  console.log("\nSearching stored missing info for each Class List row for subject RSE:");
  classListData.forEach((row, index) => {
    const missingEntry = searchMissingInfo(missingInfoMap, row.Form, row.Class, "RSE", row.RSE);
    if (missingEntry) {
      console.log(`Row ${index + 1} -> Matching group found for RSE: ${missingEntry.group}`);

      // Create a combined string "group subject teacher" e.g. "3MJPW4 RSE NEW"
      const combinedGrouping = `${missingEntry.group} ${missingEntry.subject} ${missingEntry.teacher}`;
      const subject_code = Object.hasOwnProperty.call(subCodeMap, missingEntry.sub_code) ? subCodeMap[missingEntry.sub_code] : null;

      // Add a new output record using this matching info.
      outputRows.push({
        tsss_id: row.TSSSID,
        a_year: "2025/2026",
        terms: "Term 1",
        sub_code: missingEntry.sub_code,
        subject_code: subject_code,
        taking: "TRUE",
        self_taking: "FALSE",
        grouping: combinedGrouping,
        grouping_real: combinedGrouping
      });
    } else {
      console.log(`Row ${index + 1} -> No matching group found for RSE`);
    }
  });

  // NEW: Searching stored missing info for each Class List row for subject "ENG"
  console.log("\nSearching stored missing info for each Class List row for subject ENG:");
  classListData.forEach((row, index) => {
    const missingEntry = searchMissingInfo(missingInfoMap, row.Form, row.Class, "ENG", row.ENG);
    if (missingEntry) {
      console.log(`Row ${index + 1} -> Matching group found for ENG: ${missingEntry.group}`);

      // Create a combined string "group subject teacher" e.g. "3MJPW4 ENG NEW"
      const combinedGrouping = `${missingEntry.group} ${missingEntry.subject} ${missingEntry.teacher}`;
      const subject_code = Object.hasOwnProperty.call(subCodeMap, missingEntry.sub_code) ? subCodeMap[missingEntry.sub_code] : null;

      // Add a new output record using this matching info.
      outputRows.push({
        tsss_id: row.TSSSID,
        a_year: "2025/2026",
        terms: "Term 1",
        sub_code: missingEntry.sub_code,
        subject_code: subject_code,
        taking: "TRUE",
        self_taking: "FALSE",
        grouping: combinedGrouping,
        grouping_real: combinedGrouping
      });
    } else {
      console.log(`Row ${index + 1} -> No matching group found for ENG`);
    }
  });

  // **New Section 1**
  // Store information where:
  // 1. Group has only 1 numeric digit.
  // 2. Subject is exactly "(M2)".
  const missingSpecific = [];
  missingInfoMap.forEach((entries, group) => {
    entries.forEach(info => {
      // Use regex to find all numeric digits in the group.
      const digits = info.group.match(/\d/g);
      console.log("digits: " + digits);
      console.log("info.subject:" + info.subject);
      console.log("-----------------");
      if (digits && digits.length === 1 && info.subject === "(M2)") {
        missingSpecific.push(info);
      }
    });
  });

  // Print out the stored information for verification.
  console.log("\nMissing class token information where group has only 1 number and subject is '(M2)':");
  if (missingSpecific.length > 0) {
    missingSpecific.forEach(info => {
      console.log(`Group: ${info.group}, Subject: ${info.subject}, Teacher: ${info.teacher}`);
    });
  } else {
    console.log("No missing entries found meeting the criteria.");
  }

  // **New Section 2**
  // Loop through "2024-25 Class List.xlsx" and, for each row where:
  //   - The 'Form' column value > 3, AND
  //   - The 'M1&2  MU' column value contains the string 'M2'
  // search missingSpecific to determine the teacher for that form.
  // And then, add a record with:
  //    tsss_id: that row's TSSSID,
  //    a_year: "2025/2026",
  //    terms: "Term 1",
  //    sub_code: "DMA21",
  //    subject_code: "DMATH2",
  //    taking: "TRUE",
  //    self_taking: "FALSE",
  //    grouping: form + ' (M2) ' + teacher,
  //    grouping_real: form + ' (M2) ' + teacher
  console.log("\nSearching for teacher using missingSpecific for class list rows (Form > 3 and 'M1&2  MU' contains 'M2'):");
  classListData.forEach(row => {
    const formNumber = Number(row.Form);
    if (formNumber > 3 && row["M1&2"] && row["M1&2"].toString().indexOf("M2") !== -1) {
      const match = missingSpecific.find(entry => entry.group.startsWith(String(formNumber)));
      if (match) {
        console.log(`Form ${formNumber} -> Teacher found: ${match.teacher} (from missingSpecific entry with group ${match.group})`);
        // Add record to outputRows:
        outputRows.push({
          tsss_id: row.TSSSID,
          a_year: "2025/2026",
          terms: "Term 1",
          sub_code: "DMA21",
          subject_code: "DMATH2",
          taking: "TRUE",
          self_taking: "FALSE",
          grouping: `${formNumber} (M2) ${match.teacher}`,
          grouping_real: `${formNumber} (M2) ${match.teacher}`
        });
      } else {
        console.log(`Form ${formNumber} -> No matching teacher found in missingSpecific.`);
      }
    }
  });

  // Finally, create a new workbook and worksheet with the output data.
  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.json_to_sheet(outputRows, {
    header: ["tsss_id", "a_year", "terms", "sub_code", "subject_code", "taking", "self_taking", "grouping", "grouping_real"]
  });
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");

  // Write the new workbook to an Excel file using the output file name defined in the config.
  XLSX.writeFile(newWorkbook, studentProfile);

  console.log("Excel file '" + studentProfile + "' has been generated with", outputRows.length, "rows.");

  // Print out rows that haven't been processed along with the reasons.
  if (unprocessedRows.length > 0) {
    console.log("\nThe following rows from '" + userProfileFile + "' were not processed:");
    unprocessedRows.forEach(entry => {
      console.log(`Row ${entry.rowNumber}: ${entry.reason}`);
    });
  } else {
    console.log("\nAll rows have been processed.");
  }
}

// Execute the main function
main();