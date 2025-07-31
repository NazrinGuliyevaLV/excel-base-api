const express = require('express');
const Excel = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');
const axios = require('axios');

const app = express();
app.use(express.json({ limit: '20mb' }));

const tempDir = os.tmpdir();
const handwritingApiUrl = 'https://createsignature.onrender.com/handwriting';

const randomName = (len) => {
    let name = '';
    for (let i = 0; i < len; i++) {
        name += String(parseInt(Math.random() * 10));
    }
    return name;
};

app.post('/insert-multiple-images', async (req, res) => {
    try {
        console.log('STEP 1: Request received');

        const { w, h, x, y, a, b, excelFileContent, sheet } = req.body;

        if (!excelFileContent) {
            console.log('Missing excelFileContent');
            return res.status(400).json({ error: 'Missing excelFileContent in request' });
        }

        const inputPath = path.join(tempDir, `input-${randomName(10)}.xlsx`);
        fs.writeFileSync(inputPath, Buffer.from(excelFileContent, 'base64'));
        console.log('STEP 2: Excel file written to disk:', inputPath);

        const workbook = new Excel.Workbook();
        await workbook.xlsx.readFile(inputPath);
        console.log('STEP 3: Excel file loaded');

        const worksheet = workbook.getWorksheet(sheet || 1);
        if (!worksheet) {
            console.log('Worksheet not found');
            return res.status(400).json({ error: 'Worksheet not found' });
        }
        console.log('STEP 4: Worksheet loaded:', worksheet.name);

        const names = [];
        let rowIndex = 1;
        while (true) {
            const cellValue = worksheet.getCell(`A${rowIndex}`).value;
            if (!cellValue) break;
            names.push(cellValue.toString());
            rowIndex++;
        }

        console.log('STEP 5: Names extracted from column A:', names);

        const images = [];
        for (const name of names) {
            console.log(`STEP 6: Sending '${name}' to handwriting API...`);
            const response = await axios.post(handwritingApiUrl, { fullname: name }, {
                responseType: 'arraybuffer'
            });
            const base64 = Buffer.from(response.data, 'binary').toString('base64');
            images.push(base64);
        }

        const imageSize = { width: w, height: h };
        let currentX = x;
        let currentY = y;

        for (let i = 0; i < images.length; i++) {
            const imgBuffer = Buffer.from(images[i], 'base64');
            const imageId = workbook.addImage({
                buffer: imgBuffer,
                extension: 'png',
            });

            worksheet.addImage(imageId, {
                tl: { col: currentX, row: currentY },
                ext: imageSize
            });

            console.log(`STEP 7: Image ${i + 1} added at col ${currentX}, row ${currentY}`);

            currentX += a;
            currentY += b;
        }

        const outputPath = path.join(tempDir, `output-${randomName(10)}.xlsx`);
        await workbook.xlsx.writeFile(outputPath);
        console.log('STEP 8: Output Excel saved:', outputPath);

        const fileBuffer = fs.readFileSync(outputPath);
        const base64Excel = fileBuffer.toString('base64');

        fs.unlinkSync(inputPath);
        fs.unlinkSync(outputPath);
        console.log('STEP 9: Temp files cleaned up');

        console.log('STEP 10: Returning result');
        return res.status(201).send(base64Excel);

    } catch (err) {
        console.error('ERROR:', err.message);
        console.error('STACK:', err.stack);
        return res.status(500).json({ error: 'Internal server error' });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
