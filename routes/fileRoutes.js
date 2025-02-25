
const express = require('express');
const fs = require('fs').promises;
const path = require('path');
const xlsx = require('xlsx');  

const router = express.Router();

// Helper function to get absolute file path (prevents path traversal attacks)
const getFilePath = (fileName) => path.join(__dirname, '..', 'storage', path.basename(fileName));

// Function to check if the file is an Excel file (xlsx)
const isExcelFile = (fileName) => fileName.endsWith('.xlsx');

router.get('/read', async (req, res) => {
  const fileName = req.query.fileName;
  
  if (!isExcelFile(fileName)) {
    return res.status(400).json({ error: 'Only Excel (.xlsx) files are allowed' });
  }

  try {
    const filePath = getFilePath(fileName);
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];  
    const sheet = workbook.Sheets[sheetName];

   

    
    const data = xlsx.utils.sheet_to_json(sheet, { defval: "", header: 1 }); 

    console.log("Data extracted:", data); 

    if (data.length === 0) {
      return res.status(404).json({ error: 'No data found in the Excel file' });
    }

    res.json({ content: data });
  } catch (err) {
    res.status(404).json({ error: 'File not found or invalid Excel file', details: err.message });
  }
});

router.post('/write', async (req, res) => {
  const { fileName, content } = req.body;
  
  if (!isExcelFile(fileName)) {
    return res.status(400).json({ error: 'Only Excel (.xlsx) files are allowed' });
  }

  try {
    const filePath = getFilePath(fileName);
    const worksheet = xlsx.utils.aoa_to_sheet(content);  
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    xlsx.writeFile(workbook, filePath);
    res.json({ message: 'File written successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});
// router.post('/append', async (req, res) => {
//   const { fileName, content } = req.body;

//   if (!isExcelFile(fileName)) {
//     return res.status(400).json({ error: 'Only Excel (.xlsx) files are allowed' });
//   }

//   try {
//     const filePath = getFilePath(fileName);
//     const workbook = xlsx.readFile(filePath);
//     const worksheet = workbook.Sheets[workbook.SheetNames[0]];
//     const existingData = xlsx.utils.sheet_to_json(worksheet, { header: 1 }); 
    
   
//     const updatedData = existingData.concat(content);  
    
//     const updatedWorksheet = xlsx.utils.aoa_to_sheet(updatedData); 
//     workbook.Sheets[workbook.SheetNames[0]] = updatedWorksheet;
//     xlsx.writeFile(workbook, filePath);
    
//     res.json({ message: 'Content appended successfully' });
//   } catch (err) {
//     res.status(500).json({ error: err.message });
//   }
// });
router.post('/append', async (req, res) => {
  const { fileName, content } = req.body;

  if (!isExcelFile(fileName)) {
    return res.status(400).json({ error: 'Only Excel (.xlsx) files are allowed' });
  }

  try {
    const filePath = getFilePath(fileName);
    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const existingData = xlsx.utils.sheet_to_json(worksheet, { header: 1 }); // قراءة البيانات على شكل مصفوفة من المصفوفات
    
    // تأكد أن content هي مصفوفة تحتوي على نص واحد في كل صف
    const newRow = [content];  // إذا كانت content نص واحد فقط، سنضعها في مصفوفة واحدة
    
    // دمج البيانات الجديدة مع البيانات القديمة
    const updatedData = existingData.concat(newRow);  // دمج البيانات القديمة مع الجديدة
    
    const updatedWorksheet = xlsx.utils.aoa_to_sheet(updatedData); // تحويل البيانات المدمجة إلى ورقة عمل جديدة
    workbook.Sheets[workbook.SheetNames[0]] = updatedWorksheet;
    xlsx.writeFile(workbook, filePath);
    
    res.json({ message: 'Content appended successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.delete('/delete', async (req, res) => {
  const { fileName } = req.query;

  if (!isExcelFile(fileName)) {
    return res.status(400).json({ error: 'Only Excel (.xlsx) files are allowed' });
  }

  try {
    const filePath = getFilePath(fileName);
    await fs.unlink(filePath); 
    res.json({ message: 'File deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.put('/rename', async (req, res) => {
  const { oldName, newName } = req.body;

  if (!isExcelFile(oldName) || !isExcelFile(newName)) {
    return res.status(400).json({ error: 'Only Excel (.xlsx) files are allowed' });
  }

  try {
    const oldFilePath = getFilePath(oldName);
    const newFilePath = getFilePath(newName);

    await fs.access(oldFilePath); 
    await fs.rename(oldFilePath, newFilePath); 
    res.json({ message: 'File renamed successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;