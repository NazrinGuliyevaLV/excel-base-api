const { app } = require('@azure/functions');

const Excel = require('exceljs');

const fs = require('fs');

const os = require('os');

const path = require('path');

const tempDir = os.tmpdir();
 
const randomName = (len) => {

    let name = '';

    for (let i = 0; i < len; i++) {

        name += String(parseInt(Math.random() * 10));

    };

    return name;

};
 
app.http('InsertImageXlsx', {

    methods: ['POST'],

    authLevel: 'anonymous',

    handler: async (req, context) => {

        context.log(`Http function processed request for url "${req.url}"`);
 
        try {

            const { x, y, h, w, sheet, excelFileContent, imageContent } = await req.json();
 
            if (!excelFileContent || !imageContent) {

                return {

                    status: 400,

                    jsonBody: { error: 'Missing fields in request' }

                };

            }
 
            const inputPath = path.join(tempDir, `input-${randomName(10)}.xlsx`);

            fs.writeFileSync(inputPath, Buffer.from(excelFileContent, 'base64'));
 
            const workbook = new Excel.Workbook();

            await workbook.xlsx.readFile(inputPath);

            const worksheet = workbook.getWorksheet(sheet);
 
            const imageBuffer = Buffer.from(imageContent, 'base64');

            const imageId = workbook.addImage({

                buffer: imageBuffer,

                extension: 'png'

            });
 
            worksheet.addImage(imageId, {

                tl: { col: x, row: y },

                ext: { width: w, height: h }

            });
 
            const outputPath = path.join(tempDir, `output-${randomName(10)}.xlsx`);

            await workbook.xlsx.writeFile(outputPath);
 
            const fileBuffer = fs.readFileSync(outputPath);

            const base64Excel = fileBuffer.toString('base64');
 
            fs.unlinkSync(inputPath);

            fs.unlinkSync(outputPath);
 
            return {

                status: 201,

                body: base64Excel

            };
 
        } catch (err) {

            context.log('err message::', err.message);

            context.log('err stack:', err.stack);

            return {

                status: 500,

                jsonBody: {

                    error: 'Internal server error'

                }

            };

        }

    }

});
 