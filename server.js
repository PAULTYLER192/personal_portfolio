const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const fs = require('fs');

const app = express();
const upload = multer();
const PORT = 5000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cors());

const FILE_PATH = 'contact_data.xlsx';

// Handle form submission
app.post('/submit-form', upload.none(), (req, res) => {
  const { name, email, message } = req.body;

  // Load existing workbook or create new one
  let workbook;
  if (fs.existsSync(FILE_PATH)) {
    workbook = xlsx.readFile(FILE_PATH);
  } else {
    workbook = xlsx.utils.book_new();
  }

  const sheetName = 'Contacts';
  let worksheet = workbook.Sheets[sheetName];

  // Create worksheet if it doesn't exist
  if (!worksheet) {
    worksheet = xlsx.utils.aoa_to_sheet([['Name', 'Email', 'Message', 'Date']]);
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
  }

  // Append new data
  const newRow = [name, email, message, new Date().toLocaleString()];
  const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  sheetData.push(newRow);
  const newWorksheet = xlsx.utils.aoa_to_sheet(sheetData);
  workbook.Sheets[sheetName] = newWorksheet;

  // Save the updated workbook
  xlsx.writeFile(workbook, FILE_PATH);

  res.status(200).json({ message: 'Form submitted successfully!' });
});

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
