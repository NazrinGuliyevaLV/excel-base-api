const express = require('express');
const Excel   = require('exceljs');
const path    = require('path');

const app = express();
app.use(express.json({ limit: '10mb' }));   // JSON + Base64 rahat sığsın

app.post('/add-image', async (req, res) => {
  try {
    const { imageContent, fileName = 'image.jpg' } = req.body || {};

    if (!imageContent) {
      return res.status(400).json({ error: 'imageContent missing' });
    }

    // 1) Base64 → Buffer
    const buffer = Buffer.from(imageContent, 'base64');
    const ext = path.extname(fileName).slice(1).toLowerCase() || 'jpg';

    // 2) Excel’i aç
    const wb = new Excel.Workbook();
    await wb.xlsx.readFile('Travel Form.xlsx');
    const ws = wb.getWorksheet('Sheet1') || wb.addWorksheet('Sheet1');

    // 3) Resmi ekle
    const imgId = wb.addImage({ buffer, extension: ext });
    ws.addImage(imgId, { tl: { col: 0, row: 0 }, br: { col: 2, row: 10 } });

    // 4) Kaydet ve cevapla
    const out = `TravelForm_${Date.now()}.xlsx`;
    await wb.xlsx.writeFile(out);
    res.json({ success: true, file: out });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.listen(3000, () => console.log('API → http://localhost:3000/add-image'));
