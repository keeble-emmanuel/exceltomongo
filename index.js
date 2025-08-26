// server.js

// Import required modules
const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Connect to MongoDB
// NOTE: Make sure your MongoDB server is running.
// For a local setup, this URL is generally sufficient.
const dbURI = 'mongodb://localhost:27017/excel_import_db';
mongoose.connect(dbURI)
  

// --- Multer Configuration for File Uploads ---
// Multer is a middleware for handling multipart/form-data, used for file uploads.
const storage = multer.diskStorage({
  // Set the destination directory for the uploaded files.
  // The 'uploads' folder must exist in your project root.
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  // Set the filename for the uploaded file.
  // We use a timestamp to ensure the filename is unique.
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  }
});

const upload = multer({ storage: storage });

// --- Mongoose Schema and Model ---
// Define a schema to represent the structure of your Excel data.
// This schema should match the columns in your Excel file.
// For this example, we'll assume the Excel file has 'name', 'age', and 'email' columns.
const dataSchema = new mongoose.Schema({
  name: {
    type: String,
    required: true
  },
  age: {
    type: Number,
    required: true
  },
  email: {
    type: String,
    required: true,
    unique: true // Ensure emails are unique
  }
});

// Create a Mongoose model based on the schema.
const DataModel = mongoose.model('Data', dataSchema);

// --- Express Route to Handle File Upload and Data Transfer ---
// This POST route uses Multer to handle the file upload. The 'excelFile' is the field name
// from the HTML form's input.
app.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    // Check if a file was uploaded.
    if (!req.file) {
      return res.status(400).json({ message: 'No file uploaded.' });
    }

    const filePath = req.file.path;

    // --- Process the Excel File ---
    // Read the Excel workbook from the temporary file path.
    const workbook = xlsx.readFile(filePath);
    
    // Get the first sheet name.
    const sheetName = workbook.SheetNames[0];
    
    // Get the worksheet object.
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert the worksheet data to a JSON array.
    // Each row in the Excel sheet becomes an object in the array.
    const jsonData = xlsx.utils.sheet_to_json(worksheet);

    // --- Insert Data into MongoDB ---
    // Use insertMany for an efficient bulk insert of all documents.
    // This is much faster than inserting one document at a time.
    const result = await DataModel.insertMany(jsonData);

    // --- Clean up the temporary file ---
    // Delete the uploaded file from the server's file system.
    fs.unlinkSync(filePath);

    // Send a success response.
    res.status(200).json({
      message: 'Data successfully transferred to MongoDB.',
      insertedCount: result.length
    });

  } catch (error) {
    console.error('Error during data transfer:', error);
    // Send a detailed error response.
    res.status(500).json({
      message: 'An error occurred during the data transfer.',
      error: error.message
    });
  }
});

// --- Simple HTML Form for Uploading ---
// Serve the HTML file to the client.
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Start the server.
app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
