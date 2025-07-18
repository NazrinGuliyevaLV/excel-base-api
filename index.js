const express = require('express');
const Excel = require('exceljs');
const fs = require('fs');
const path = require('path');
const bodyParser = require('body-parser')



const app = express(); 


app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

app.use(bodyParser.json());
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

    const worksheet = workbook.getWorksheet('form') || workbook.addWorksheet('form');
 
    const ext = path.extname(fileName).slice(1) || 'jpeg';
    const imageBuffer = Buffer.from(imageContent, 'base64');

   

    const imageId = workbook.addImage({
  buffer: imageBuffer,
  extension: ext
});
 
worksheet.getRow(9).height = 80;
worksheet.getColumn(2).width = 25;
 
worksheet.addImage(imageId, {
  tl: { col: 1, row: 8 },  
  ext: { width: 150, height: 100 }
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
