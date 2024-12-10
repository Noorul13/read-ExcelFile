const express = require("express");
const xlsx = require("xlsx");
const path = require("path");

const app = express();
const PORT = 3000;

app.use(express.static(path.join(__dirname, "public")));

// Function to read Excel data
function readExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
  const sheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(sheet);
  return jsonData;
}

// Endpoint to fetch Excel data
app.get("/data", (req, res) => {
  const filePath = path.join(__dirname, "data.xlsx");
  try {
    const data = readExcel(filePath);
    res.json(data); // Send the Excel data as JSON
  } catch (error) {
    res.status(500).send("Error reading Excel file: " + error.message);
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
