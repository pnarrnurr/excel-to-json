const http = require('http');
const express = require('express');
const xlsx = require('xlsx');
const fs = require('fs');
const app = express();

const hostname = '127.0.0.1';
const port = 3000;

app.get("/", function (req, res) {
  res.statusCode = 200;
  // res.setHeader('Content-Type', 'text/html');

  res.writeHead(200, { 'Content-Type': 'text/html' });
  res.write('<form action="/uploadExcel" method="POST" enctype="multipart/form-data">');
  res.write('<input type="file" name="excelFile" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"><br>');
  res.write('<input type="submit">');
  res.write('</form>');

  res.end('Hello World');


  convertExcelFileToJsonUsingXlsx();
});

app.post("/uploadExcel", function (req, res) {
  console.log(res)
  console.log("burada");
})

function convertExcelFileToJsonUsingXlsx() {
  // Read the file using pathname
  const file = xlsx.readFile('./Data.xlsx');
  // Grab the sheet info from the file
  const sheetNames = file.SheetNames;
  const totalSheets = sheetNames.length;
  // Variable to store our data 
  let parsedData = [];
  // Loop through sheets
  for (let i = 0; i < totalSheets; i++) {
    // Convert to json using xlsx
    const tempData = xlsx.utils.sheet_to_json(file.Sheets[sheetNames[i]], { defval: "" });
    // console.log(file.Sheets[sheetNames[i]]);
    // Skip header row which is the colum names
    // tempData.shift();
    // Add the sheet's json to our data array
    parsedData.push(...tempData);
  }
  // call a function to save the data in a json file
  generateJSONFile(parsedData);
}

function generateJSONFile(data) {
  try {
    fs.writeFileSync('data.json', JSON.stringify(data))
  } catch (err) {
    console.error(err)
  }
}

app.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});