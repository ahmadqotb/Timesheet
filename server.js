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

const uploadsDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');

[uploadsDir, outputDir].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

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

        // Generate individual PDFs into the unique run folder
        await generator.generatePDFs(runDir, (data) => {
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        });

        // Create the ZIP archive
        const zipName = `Timesheets_${timestamp}.zip`;
        const zipPath = path.join(outputDir, zipName);
        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip');

        await new Promise((resolve, reject) => {
            output.on('close', resolve);
            archive.on('error', reject);
            archive.pipe(output);
            archive.directory(runDir, false);
            archive.finalize();
        });

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
    const fileName = req.query.fileName;
    const filePath = path.join(outputDir, fileName);
    if (fs.existsSync(filePath)) res.download(filePath);
    else res.status(404).send('File not found');
});

// Cleanup: Delete files older than 30 minutes
setInterval(() => {
    const now = Date.now();
    fs.readdirSync(outputDir).forEach(f => {
        const p = path.join(outputDir, f);
        if (now - fs.statSync(p).mtimeMs > 1800000) fs.rmSync(p, { recursive: true, force: true });
    });
}, 600000);

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
