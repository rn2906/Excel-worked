// const xlsx = require("xlsx");
// const fs = require("fs");

// // Load the Excel file
// const excelFilePath = "convert.xlsx";
// const workbook = xlsx.readFile(excelFilePath);
// const sheetName = workbook.SheetNames[0];
// const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

// // Load the JSON file
// const jsonFilePath = "Extracted_Text.json";
// const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, "utf8"));

// // Config file
// const config = {
//   mappings: {
//     "Security ID": "Security_ID",
//     ISIN: "ISIN",
//     "Coupon Rate": "Coupon_Rate",
//     "Maturity Date": "Maturity_Date",
//     "Issue Date": "Issue_Date",
//     "Transaction ID": "Transaction_ID",
//     "Transaction Number": "Transaction_Number",
//     Activity: "Activity",
//     Quantity: "Quantity",
//     Price: "Price",
//     Principle: "Principle",
//     Interest: "Interest",
//     "Net Amount": "Net_Amount",
//     "Trade Date": "Trade_Date",
//     "Settlement Date": "Settlement_Date",
//     "Account Number": "Account_Number",
//   },
// };

// // Function to compare records and identify differences
// function compareRecords(excelRow, jsonRecord, mappings) {
//   const differences = {};
//   for (const [excelCol, jsonKey] of Object.entries(mappings)) {
//     const excelValue = excelRow[excelCol] || null;
//     const jsonValue = jsonRecord[jsonKey] || null;
//     if (excelValue !== jsonValue) {
//       differences[excelCol] = { Excel: excelValue, JSON: jsonValue };
//     }
//   }
//   return differences;
// }

// // Log JSON data sample to verify structure
// console.log("JSON Data Sample:", JSON.stringify(jsonData[0], null, 2));

// // Extract unique account numbers from JSON data
// const accountNumbers = new Set(jsonData.map((record) => record.Account_Number));
// console.log("Extracted Account Numbers from JSON Data:", [...accountNumbers]);

// // Filter Excel data to keep only records with those account numbers
// const filteredExcelData = excelData.filter((row) =>
//   accountNumbers.has(row["Account_Number"])
// );

// // Log filtered Excel data to verify structure
// console.log("Filtered Excel Data:", JSON.stringify(filteredExcelData, null, 2));

// // Find and print differences
// filteredExcelData.forEach((excelRow) => {
//   const transactionId = excelRow["Transaction ID"];
//   jsonData.forEach((jsonRecord) => {
//     if (jsonRecord.Transaction_ID === transactionId) {
//       const differences = compareRecords(excelRow, jsonRecord, config.mappings);
//       if (Object.keys(differences).length > 0) {
//         console.log(`Differences found for Transaction ID ${transactionId}:`);
//         for (const [col, diff] of Object.entries(differences)) {
//           console.log(`  ${col}: Excel = ${diff.Excel}, JSON = ${diff.JSON}`);
//         }
//       } else {
//         console.log(`No differences found for Transaction ID ${transactionId}`);
//       }
//     }
//   });
// });

// console.log("Comparison complete.");
//////////////////////////////////////////////////////////////////////////////////
// const xlsx = require("xlsx");
// const fs = require("fs");

// // Load the Excel file
// const excelFilePath = "convert.xlsx";
// const workbook = xlsx.readFile(excelFilePath);
// const sheetName = workbook.SheetNames[0];
// const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

// // Load the JSON file
// const jsonFilePath = "Extracted_Text.json";
// const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, "utf8"));

// // Config file
// const config = {
//   mappings: {
//     "Security ID": "Security_ID",
//     ISIN: "ISIN",
//     "Coupon Rate": "Coupon_Rate",
//     "Maturity Date": "Maturity_Date",
//     "Issue Date": "Issue_Date",
//     "Transaction ID": "Transaction_ID",
//     "Transaction Number": "Transaction_Number",
//     Activity: "Activity",
//     Quantity: "Quantity",
//     Price: "Price",
//     Principle: "Principle",
//     Interest: "Interest",
//     "Net Amount": "Net_Amount",
//     "Trade Date": "Trade_Date",
//     "Settlement Date": "Settlement_Date",
//     "Account Number": "Account_Number",
//   },
// };

// // Function to compare records and identify differences
// function compareRecords(excelRow, jsonRecord, mappings) {
//   const differences = {};
//   for (const [excelCol, jsonKey] of Object.entries(mappings)) {
//     const excelValue = excelRow[excelCol] || null;
//     const jsonValue = jsonRecord[jsonKey] || null;
//     if (excelValue !== jsonValue) {
//       differences[excelCol] = { Excel: excelValue, JSON: jsonValue };
//     }
//   }
//   return differences;
// }

// // Log JSON data sample to verify structure
// console.log("JSON Data Sample:", JSON.stringify(jsonData[0], null, 2));

// // Extract unique account numbers from JSON data
// const accountNumbers = new Set(jsonData.map((record) => record.Account_Number));
// console.log("Extracted Account Numbers from JSON Data:", [...accountNumbers]);

// // Filter Excel data to keep only records with those account numbers
// const filteredExcelData = excelData.filter((row) =>
//   accountNumbers.has(row["Account Number"])
// );

// // Log filtered Excel data to verify structure
// console.log("Filtered Excel Data:", JSON.stringify(filteredExcelData, null, 2));

// // Function to print differences in both files
// function printDifferences(excelRow, jsonRecord, differences) {
//   console.log(`Transaction ID ${excelRow["Transaction ID"]}:`);
//   console.log("Excel Record:", JSON.stringify(excelRow, null, 2));
//   console.log("JSON Record:", JSON.stringify(jsonRecord, null, 2));
//   console.log("Differences:");
//   for (const [col, diff] of Object.entries(differences)) {
//     console.log(`  ${col}: Excel = ${diff.Excel}, JSON = ${diff.JSON}`);
//   }
// }

// // Find and print differences
// filteredExcelData.forEach((excelRow) => {
//   const transactionId = excelRow["Transaction ID"];
//   jsonData.forEach((jsonRecord) => {
//     if (jsonRecord.Transaction_ID === transactionId) {
//       const differences = compareRecords(excelRow, jsonRecord, config.mappings);
//       if (Object.keys(differences).length > 0) {
//         printDifferences(excelRow, jsonRecord, differences);
//       } else {
//         console.log(`No differences found for Transaction ID ${transactionId}`);
//       }
//     }
//   });
// });

// console.log("Comparison complete.");

//////////////////////////////////////////////////////////////////////////////////////////////////////////////

// const xlsx = require("xlsx");
// const fs = require("fs");

// // Load the Excel file
// const excelFilePath = "convert.xlsx";
// const workbook = xlsx.readFile(excelFilePath);
// const sheetName = workbook.SheetNames[0];
// const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

// // Load the JSON file
// const jsonFilePath = "Extracted_Text.json";
// const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, "utf8"));

// // Config file
// const config = {
//   mappings: {
//     Security_ID: "Security_ID",
//     ISIN: "ISIN",
//     Coupon_Rate: "Coupon_Rate",
//     Maturity_Date: "Maturity_Date",
//     Issue_Date: "Issue_Date",
//     Transaction_ID: "Transaction_ID",
//     Transaction_Number: "Transaction_Number",
//     Activity: "Activity",
//     Quantity: "Quantity",
//     Price: "Price",
//     Principle: "Principle",
//     Interest: "Interest",
//     Net_Amount: "Net_Amount",
//     Trade_Date: "Trade_Date",
//     Settlement_Date: "Settlement_Date",
//     Account_Number: "Account_Number",
//   },
// };

// // Function to compare records and identify differences
// function compareRecords(excelRow, jsonRecord, mappings) {
//   const differences = {};
//   for (const [excelCol, jsonKey] of Object.entries(mappings)) {
//     const excelValue = String(excelRow[excelCol]).trim();
//     const jsonValue = String(jsonRecord[jsonKey]).trim();
//     if (excelValue != jsonValue) {
//       differences[excelCol] = { Excel: excelValue, JSON: jsonValue };
//     }
//   }
//   return differences;
// }

// // Function to print differences in both files
// function printDifferences(excelRow, jsonRecord, differences) {
//   console.log(`Transaction ID ${excelRow["Transaction_ID"]}:`);
//   console.log("Excel Record:", JSON.stringify(excelRow, null, 2));
//   console.log("JSON Record:", JSON.stringify(jsonRecord, null, 2));
//   console.log("Differences:");
//   for (const [col, diff] of Object.entries(differences)) {
//     console.log(`  ${col}: Excel = "${diff.Excel}", JSON = "${diff.JSON}"`);
//   }
// }

// // Log JSON data sample to verify structure
// console.log("JSON Data Sample:", JSON.stringify(jsonData[0], null, 2));

// // Extract unique account numbers from JSON data
// const accountNumbers = new Set(
//   jsonData.map((record) => String(record.Account_Number).trim())
// );
// console.log("Extracted Account Numbers from JSON Data:", [...accountNumbers]);

// // Log Excel data keys to verify column names
// console.log("Excel Data Keys:", Object.keys(excelData[0]));

// // Inspect the Excel data to verify account number field
// const excelAccountNumbers = new Set(
//   excelData.map((row) => {
//     if (row["Account_Number"] !== undefined) {
//       return String(row["Account_Number"]).trim();
//     }
//     return "undefined";
//   })
// );
// console.log("Excel Account Numbers:", [...excelAccountNumbers]);

// // Filter Excel data to keep only records with those account numbers
// const filteredExcelData = excelData.filter((row) =>
//   accountNumbers.has(String(row["Account_Number"]).trim())
// );

// // Log filtered Excel data to verify structure
// console.log("Filtered Excel Data:", JSON.stringify(filteredExcelData, null, 2));

// // Find and print differences
// filteredExcelData.forEach((excelRow) => {
//   const transactionId = excelRow["Transaction_ID"];
//   jsonData.forEach((jsonRecord) => {
//     if (jsonRecord.Transaction_ID == transactionId) {
//       const differences = compareRecords(excelRow, jsonRecord, config.mappings);
//       if (Object.keys(differences).length > 0) {
//         printDifferences(excelRow, jsonRecord, differences);
//       } else {
//         console.log(`No differences found for Transaction ID ${transactionId}`);
//       }
//     }
//   });
// });

// console.log("Comparison complete.");

///////////////////////////////////////////////////////////////////////////////////////////////////////////

// const xlsx = require("xlsx");
// const fs = require("fs");

// // Load the Excel file
// const excelFilePath = "convert.xlsx";
// const workbook = xlsx.readFile(excelFilePath);
// const sheetName = workbook.SheetNames[0];
// const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

// // Load the JSON file
// const jsonFilePath = "Extracted_Text.json";
// const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, "utf8"));

// // Config file
// const config = {
//   mappings: {
//     Security_ID: "Security_ID",
//     ISIN: "ISIN",
//     Coupon_Rate: "Coupon_Rate",
//     Maturity_Date: "Maturity_Date",
//     Issue_Date: "Issue_Date",
//     Transaction_ID: "Transaction_ID",
//     Transaction_Number: "Transaction_Number",
//     Activity: "Activity",
//     Quantity: "Quantity",
//     Price: "Price",
//     Principle: "Principle",
//     Interest: "Interest",
//     Net_Amount: "Net_Amount",
//     Trade_Date: "Trade_Date",
//     Settlement_Date: "Settlement_Date",
//     Account_Number: "Account_Number",
//   },
//   dateFields: ["Maturity_Date", "Issue_Date", "Trade_Date", "Settlement_Date"], // Explicitly specify date fields
// };

// // Function to normalize dates
// function normalizeDate(value) {
//   if (typeof value === "number") {
//     // Excel date serial numbers are stored as numbers
//     const date = new Date((value - 25567) * 86400 * 1000);
//     return date.toISOString().split("T")[0]; // Format as YYYY-MM-DD
//   }
//   if (typeof value === "string") {
//     const date = new Date(value);
//     return date.toISOString().split("T")[0]; // Format as YYYY-MM-DD
//   }
//   return value;
// }

// // Function to normalize values based on type
// function normalizeValue(value, isDateField) {
//   if (isDateField) {
//     return normalizeDate(value);
//   }
//   return value; // Return as-is for non-date fields
// }

// // Function to compare records and identify differences
// function compareRecords(excelRow, jsonRecord, mappings, dateFields) {
//   const differences = {};
//   for (const [excelCol, jsonKey] of Object.entries(mappings)) {
//     const isDateField = dateFields.includes(jsonKey);
//     const excelValue = normalizeValue(excelRow[excelCol], isDateField);
//     const jsonValue = normalizeValue(jsonRecord[jsonKey], isDateField);
//     if (excelValue !== jsonValue) {
//       differences[excelCol] = { Excel: excelValue, JSON: jsonValue };
//     }
//   }
//   return differences;
// }

// // Function to print differences in both files
// function printDifferences(excelRow, jsonRecord, differences) {
//   console.log(`Transaction ID ${excelRow["Transaction_ID"]}:`);
//   console.log("Excel Record:", JSON.stringify(excelRow, null, 2));
//   console.log("JSON Record:", JSON.stringify(jsonRecord, null, 2));
//   console.log("Differences:");
//   for (const [col, diff] of Object.entries(differences)) {
//     console.log(`  ${col}: Excel = "${diff.Excel}", JSON = "${diff.JSON}"`);
//   }
// }

// // Log JSON data sample to verify structure
// console.log("JSON Data Sample:", JSON.stringify(jsonData[0], null, 2));

// // Extract unique account numbers from JSON data
// const accountNumbers = new Set(
//   jsonData.map((record) => String(record.Account_Number).trim())
// );
// console.log("Extracted Account Numbers from JSON Data:", [...accountNumbers]);

// // Log Excel data keys to verify column names
// console.log("Excel Data Keys:", Object.keys(excelData[0]));

// // Inspect the Excel data to verify account number field
// const excelAccountNumbers = new Set(
//   excelData.map((row) => {
//     if (row["Account_Number"] !== undefined) {
//       return String(row["Account_Number"]).trim();
//     }
//     return "undefined";
//   })
// );
// console.log("Excel Account Numbers:", [...excelAccountNumbers]);

// // Filter Excel data to keep only records with those account numbers
// const filteredExcelData = excelData.filter((row) =>
//   accountNumbers.has(String(row["Account_Number"]).trim())
// );

// // Log filtered Excel data to verify structure
// console.log("Filtered Excel Data:", JSON.stringify(filteredExcelData, null, 2));

// // Find and print differences
// filteredExcelData.forEach((excelRow) => {
//   const transactionId = excelRow["Transaction_ID"];
//   jsonData.forEach((jsonRecord) => {
//     if (jsonRecord.Transaction_ID == transactionId) {
//       const differences = compareRecords(
//         excelRow,
//         jsonRecord,
//         config.mappings,
//         config.dateFields
//       );
//       if (Object.keys(differences).length > 0) {
//         printDifferences(excelRow, jsonRecord, differences);
//       } else {
//         console.log(`No differences found for Transaction ID ${transactionId}`);
//       }
//     }
//   });
// });

// console.log("Comparison complete.");

////////////////////////////////////////////////////////////////////////////////////////////
const xlsx = require("xlsx");
const fs = require("fs");

// Load the Excel file
const excelFilePath = "convert.xlsx";
const workbook = xlsx.readFile(excelFilePath);
const sheetName = workbook.SheetNames[0];
const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

// Load the JSON file
const jsonFilePath = "Extracted_Text.json";
const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, "utf8"));

// Config file
const config = {
  mappings: {
    Security_ID: "Security_ID",
    ISIN: "ISIN",
    Coupon_Rate: "Coupon_Rate",
    Maturity_Date: "Maturity_Date",
    Issue_Date: "Issue_Date",
    Transaction_ID: "Transaction_ID",
    Transaction_Number: "Transaction_Number",
    Activity: "Activity",
    Quantity: "Quantity",

    Price: "Price",
    Principle: "Principle",
    Interest: "Interest",
    Net_Amount: "Net_Amount",
    Trade_Date: "Trade_Date",
    Settlement_Date: "Settlement_Date",
    Account_Number: "Account_Number",
  },
  dateFields: ["Maturity_Date", "Issue_Date", "Trade_Date", "Settlement_Date"], // Explicitly specify date fields
};

// Function to normalize dates
function normalizeDate(value) {
  if (typeof value === "number") {
    // Convert Excel date serial numbers to JavaScript Date
    const excelBaseDate = new Date(Date.UTC(1899, 11, 30)); // Excel starts from Dec 30, 1899
    const date = new Date(excelBaseDate.getTime() + value * 86400 * 1000);
    return date.toISOString().split("T")[0]; // Format as YYYY-MM-DD
  }
  if (typeof value === "string") {
    // Handle string dates in various formats
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toISOString().split("T")[0]; // Format as YYYY-MM-DD
    }
  }
  return value; // Return as-is for non-date values
}

// Function to normalize values based on type
function normalizeValue(value, isDateField) {
  if (isDateField) {
    return normalizeDate(value);
  }
  return value; // Return as-is for non-date fields
}

// Function to compare records and identify differences
function compareRecords(excelRow, jsonRecord, mappings, dateFields) {
  const differences = {};
  for (const [excelCol, jsonKey] of Object.entries(mappings)) {
    const isDateField = dateFields.includes(jsonKey);
    const excelValue = normalizeValue(excelRow[excelCol], isDateField);
    const jsonValue = normalizeValue(jsonRecord[jsonKey], isDateField);
    if (excelValue !== jsonValue) {
      differences[excelCol] = { Excel: excelValue, JSON: jsonValue };
    }
  }
  return differences;
}

// Function to print differences in both files
function printDifferences(excelRow, jsonRecord, differences) {
  //   console.log(`Transaction ID ${excelRow["Transaction_ID"]}:`);
  //   console.log("Excel Record:", JSON.stringify(excelRow, null, 2));
  //   console.log("JSON Record:", JSON.stringify(jsonRecord, null, 2));
  console.log("Differences:");
  for (const [col, diff] of Object.entries(differences)) {
    console.log(`  ${col}: Excel = "${diff.Excel}", JSON = "${diff.JSON}"`);
  }
}

// Log JSON data sample to verify structure
console.log("JSON Data Sample:", JSON.stringify(jsonData[0], null, 2));

// Extract unique account numbers from JSON data
const accountNumbers = new Set(
  jsonData.map((record) => String(record.Account_Number).trim())
);
console.log("Extracted Account Numbers from JSON Data:", [...accountNumbers]);

// Log Excel data keys to verify column names
console.log("Excel Data Keys:", Object.keys(excelData[0]));

// Inspect the Excel data to verify account number field
const excelAccountNumbers = new Set(
  excelData.map((row) => {
    if (row["Account_Number"] !== undefined) {
      return String(row["Account_Number"]).trim();
    }
    return "undefined";
  })
);
console.log("Excel Account Numbers:", [...excelAccountNumbers]);

// Filter Excel data to keep only records with those account numbers
const filteredExcelData = excelData.filter((row) =>
  accountNumbers.has(String(row["Account_Number"]).trim())
);

// Log filtered Excel data to verify structure
console.log("Filtered Excel Data:", JSON.stringify(filteredExcelData, null, 2));

// Find and print differences
filteredExcelData.forEach((excelRow) => {
  const transactionId = excelRow["Transaction_ID"];
  jsonData.forEach((jsonRecord) => {
    if (jsonRecord.Transaction_ID == transactionId) {
      const differences = compareRecords(
        excelRow,
        jsonRecord,
        config.mappings,
        config.dateFields
      );
      if (Object.keys(differences).length > 0) {
        printDifferences(excelRow, jsonRecord, differences);
      } else {
        console.log(`No differences found for Transaction ID ${transactionId}`);
      }
    }
  });
});

console.log("Comparison complete.");
