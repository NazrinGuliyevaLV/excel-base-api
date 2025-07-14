const express = require('express');
const Excel = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express(); 

app.post('/add-image', async (req, res) => {
  try {
    const { excelFileContent, imageContent, fileName } = req.body;

    if (!excelFileContent || !imageContent || !fileName) {
      return res.status(400).json({ error: 'Missing fields in request' });
    }
 
    const tempExcelPath = path.join(__dirname, 'temp-input.xlsx');
    fs.writeFileSync(tempExcelPath, Buffer.from(excelFileContent, 'base64'));
 
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(tempExcelPath);

    const worksheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
    const ext = path.extname(fileName).slice(1) || 'jpeg';
    const imageBuffer = Buffer.from(imageContent, 'base64');

    const imageId = workbook.addImage({
      buffer: imageBuffer,
      extension: ext
    });

    worksheet.addImage(imageId, {
      tl: { col: 0, row: 0 },
      br: { col: 3, row: 10 }
    });
 
    const outputPath = path.join(__dirname, `output_${Date.now()}.xlsx`);
    await workbook.xlsx.writeFile(outputPath);

    res.json({ success: true, savedFile: path.basename(outputPath) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(3000, () => console.log('API listening on http://localhost:3000'));
