const express = require('express');
const multer = require('multer');
const PDFExtract = require('pdf.js-extract').PDFExtract;
const excel = require('exceljs');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

const pdfExtract = new PDFExtract();  // Correct initialization

app.post('/pdf-to-excel', upload.single('pdf'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No PDF file uploaded' });
    }

    // Use pdfExtract with promises:
    const buffer = fs.readFileSync(req.file);
    // const data = await pdfExtract.extractBuffer(req.file.buffer);
    const data = pdfExtract.extractBuffer(buffer, options, (err, data) => {
        if (err) return console.log(err);
        console.log(data);

      });
    console.log(JSON.stringify(data, null, 2)); // Inspect the data structure!!!

    const rows = [];

    // ***CRITICAL: ADAPT THIS PART BASED ON YOUR PDF'S DATA STRUCTURE***
    // Example (grouping by y-coordinate - likely incorrect for your PDF):
    data.pages.forEach(page => {
      const pageRows = {};
      page.content.forEach(item => {
        if (!pageRows[item.y]) {
          pageRows[item.y] = [];
        }
        pageRows[item.y].push(item.str);
      });
      Object.values(pageRows).forEach(row => {
        rows.push(row);
      });
    });

    console.log("Extracted Rows:", rows); // Check the extracted rows

    const workbook = new excel.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    worksheet.addRows(rows);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=output.xlsx');

    const excelBuffer = await workbook.xlsx.writeBuffer();
    res.send(excelBuffer);


  } catch (error) {
    console.error("General Error:", error);
    res.status(500).json({ error: 'An error occurred during conversion' });
  }
});

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});