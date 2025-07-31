// const express = require('express');
// const Excel = require('exceljs');
// const fs = require('fs');
// const os = require('os');
// const path = require('path');

// const app = express();
// app.use(express.json({ limit: '10mb' }));  

// const tempDir = os.tmpdir();

// const randomName = (len) => {
//     let name = '';
//     for (let i = 0; i < len; i++) {
//         name += String(parseInt(Math.random() * 10));
//     }
//     return name;
// };

// app.post('/insert-image', async (req, res) => {
//     try {
//         const { x, y,a,b, h, w, sheet, excelFileContent, imageContent } = req.body;

//         if (!excelFileContent || !imageContent) {
//             return res.status(400).json({ error: 'Missing fields in request' });
//         }

//         const inputPath = path.join(tempDir, `input-${randomName(10)}.xlsx`);
//         fs.writeFileSync(inputPath, Buffer.from(excelFileContent, 'base64'));

//         const workbook = new Excel.Workbook();
//         await workbook.xlsx.readFile(inputPath);

//         const worksheet = workbook.getWorksheet(sheet);
//         const imageBuffer = Buffer.from(imageContent, 'base64');

//         const imageId = workbook.addImage({
//             buffer: imageBuffer,
//             extension: 'png',
//         });

//         worksheet.addImage(imageId, {
//             tl: { col: x, row: y },
//             ext: { width: w, height: h },
//         });

//         const outputPath = path.join(tempDir, `output-${randomName(10)}.xlsx`);
//         await workbook.xlsx.writeFile(outputPath);

//         const fileBuffer = fs.readFileSync(outputPath);
//         const base64Excel = fileBuffer.toString('base64');

//         fs.unlinkSync(inputPath);
//         fs.unlinkSync(outputPath);

//         return res.status(201).send(base64Excel);
//     } catch (err) {
//         console.error('Error:', err.message);
//         console.error('Stack:', err.stack);
//         return res.status(500).json({ error: 'Internal server error' });
//     }
// });

// const PORT = process.env.PORT || 3000;
// app.listen(PORT, () => {
//     console.log(`Server running on port ${PORT}`);
// });




const express = require('express');
const Excel = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');
const axios = require('axios');

const app = express();
app.use(express.json({ limit: '20mb' }));

const tempDir = os.tmpdir();
const handwritingApiUrl = 'https://2e1b0b71c7c0.ngrok-free.app/api/Signatures'; 

const randomName = (len) => {
    let name = '';
    for (let i = 0; i < len; i++) {
        name += String(parseInt(Math.random() * 10));
    }
    return name;
};

app.post('/insert-multiple-images', async (req, res) => {
    try {
        const { excelFileContent, sheet } = req.body;

        if (!excelFileContent) {
            return res.status(400).json({ error: 'Missing excelFileContent in request' });
        }

        const inputPath = path.join(tempDir, `input-${randomName(10)}.xlsx`);
        fs.writeFileSync(inputPath, Buffer.from(excelFileContent, 'base64'));

        const workbook = new Excel.Workbook();
        await workbook.xlsx.readFile(inputPath);
        const worksheet = workbook.getWorksheet(sheet || 1);  

        // Adlari oxu, meselen A1
        const names = [];
        let rowIndex = 1;
        while (true) {
            const cellValue = worksheet.getCell(`A${rowIndex}`).value;
            if (!cellValue) break;
            names.push(cellValue.toString());
            rowIndex++;
        }

        // her adın png -si
        const images = [];
        for (const name of names) {
            const response = await axios.post(handwritingApiUrl, { text: name });
            images.push(response.data);  
        }

        // png -ni excele elave ele
        const imageSize = { width: 200, height: 50 }; 
        let startRow = 1;
        for (let i = 0; i < images.length; i++) {
            const imgBuffer = Buffer.from(images[i], 'base64');
            const imageId = workbook.addImage({
                buffer: imgBuffer,
                extension: 'png',
            });

            worksheet.addImage(imageId, {
                tl: { col: 2, row: startRow }, //meselen b sutunundan alt alta
                ext: imageSize
            });

            startRow += 2; 
        }

        const outputPath = path.join(tempDir, `output-${randomName(10)}.xlsx`);
        await workbook.xlsx.writeFile(outputPath);
        const fileBuffer = fs.readFileSync(outputPath);
        const base64Excel = fileBuffer.toString('base64');

        fs.unlinkSync(inputPath);
        fs.unlinkSync(outputPath);

        return res.status(201).send(base64Excel);

    } catch (err) {
        console.error('Error:', err.message);
        return res.status(500).json({ error: 'Internal server error' });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
