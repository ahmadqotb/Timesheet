const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const TimesheetGenerator = require('./generator');
const DataValidator = require('./dataValidator');
const ProjectSummaryGenerator = require('./projectSummary');

const app = express();
const PORT = process.env.PORT || 3000;

const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

const upload = multer({ dest: 'uploads/' });
app.use(express.static(__dirname));

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

app.post('/generate', upload.single('excel'), async (req, res) => {
    let runDir = '';
    try {
        const { month, year } = req.body;
        const satFri = req.body.satFriEmployees ? JSON.parse(req.body.satFriEmployees) : [];
        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');

        const timestamp = Date.now();
        runDir = path.join(outputDir, `run-${timestamp}`);
        fs.mkdirSync(runDir, { recursive: true });

        const generator = new TimesheetGenerator(req.file.path, month, year, satFri);
        const count = await generator.processExcel();

        await generator.generatePDFs(runDir, (data) => {
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        });

        const zipName = `Timesheets_${timestamp}.zip`;
        const zipPath = path.join(outputDir, zipName);
        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip');

        archive.pipe(output);
        archive.directory(runDir, false);
        await archive.finalize();

        res.write(`data: ${JSON.stringify({ progress: 100, status: 'Complete!', complete: true, count, zipFile: zipName })}\n\n`);
        res.end();
    } catch (e) {
        res.write(`data: ${JSON.stringify({ error: e.message })}\n\n`);
        res.end();
    } finally {
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        if (runDir && fs.existsSync(runDir)) fs.rmSync(runDir, { recursive: true, force: true });
    }
});

app.get(['/download-all', '/download-food-allowance'], (req, res) => {
    const filePath = path.join(outputDir, req.query.fileName);
    if (fs.existsSync(filePath)) res.download(filePath);
    else res.status(404).send('File not found');
});

app.listen(PORT, () => console.log(`Server started on ${PORT}`));
