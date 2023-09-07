const XLSX = require('xlsx');

// Function to read Excel data from all rows
function readExcelData(filePath) {
  const workbook = XLSX.readFile(filePath);
  const worksheet = workbook.Sheets['Arkusz1'];
  const allRows = [];
  const startRow = 1; // Assuming data starts from row 1
  const endRow = 100; // You can adjust this to read up to a specific row

  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber++) {
    const cellA = worksheet['A' + rowNumber];
    const cellB = worksheet['B' + rowNumber];
    if (!cellA || !cellB) {
      break;
    }

    let jsonData;
  try {
      jsonData = JSON.parse(cellB.v);
    } catch (error) {
      console.error(`Error parsing JSON in row ${rowNumber}:`, error);
      jsonData = null;
    }

    // Add row data to the array
    allRows.push({ cellA: cellA.v, jsonData });
  }

  return allRows;
}

// Function to convert JSON data to HTML
function convertJsonToHtml(jsonData) {
  let html = "";
  jsonData.sections.forEach((section) => {
    section.items.forEach((item) => {
      if (item.type === 'TEXT') {
        html += item.content;
      }
    });
  });
  html = html.replace(/<p>Specjalizujemy się w hurtowej oraz detalicznej sprzedaży materiałów ściernych\.<\/p>[\s\S]+?<ul>[\s\S]+?<\/ul>/, '');

  return html;
}

const inputFilePath = 'offers.xlsx';

// Read data from all rows in the Excel file
const excelData = readExcelData(inputFilePath);

// Process each row of data
excelData.forEach((rowData, index) => {
  const htmlData = convertJsonToHtml(rowData.jsonData);
  console.log(`HTML data for Row ${index + 1}:`, htmlData);
  // You can save this HTML data to a file or perform any other desired actions
});

function saveHtmlToExcel(filePath, outputFilePath, htmlDataArray) {
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets['Arkusz1'];
  
    // Iterate through the rows and set HTML data in Column B
    htmlDataArray.forEach((htmlData, index) => {
      const rowIndex = index + 1; // Adjust for 1-based index and header row
      const cellB = `B${rowIndex}`;
  
      // Set the HTML content in Column B
      worksheet[cellB] = { t: 's', v: htmlData };
    });
  
    XLSX.writeFile(workbook, outputFilePath);
  }


  
  // Process each row of data and extract HTML
  const htmlDataArray = excelData.map((rowData) => {
    return convertJsonToHtml(rowData.jsonData);
  });
  
  // Define the output Excel file path
  const outputFilePath = 'finished.xlsx';
  
  // Save the HTML data to the Excel file in Column B
  saveHtmlToExcel(inputFilePath, outputFilePath, htmlDataArray);
  
  console.log(`HTML data saved to Excel file: ${outputFilePath}`);
