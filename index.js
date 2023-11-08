// Import the required libraries
const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');
const dJSON = require('dirty-json');

// Function to convert JSON to Excel
function jsonToExcel(inputFile, outputDirectory) {
  // Read the JSON file
  const rawData = fs.readFileSync(inputFile);
  let data;

  try {
    data = dJSON.parse(rawData);
  } catch (error) {
    console.log(`Skipping ${inputFile}`, error);
    return;
  }

  // Convert JSON to Excel
  const worksheet = XLSX.utils.json_to_sheet(data.options);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'shared-options');

  // Create output file name
  const outputFileName = path.basename(inputFile, path.extname(inputFile)) + '.xlsx';

  const outputDir = outputDirectory || 'output';
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
  }
  const outputFilePath = path.join(outputDir, outputFileName);  
  // Write to file
  XLSX.writeFile(workbook, outputFilePath);

  console.log(`Converted ${inputFile} to ${outputFilePath}`);
}

// Function to convert all JSON files in a directory
function convertAllJsonFiles(directory) {
  fs.readdirSync(directory).forEach(file => {
    if (path.extname(file) === '.json') {
      jsonToExcel(path.join(directory, file));
    }
  });
}

// Call the function with your directory as parameter
convertAllJsonFiles('json-files');