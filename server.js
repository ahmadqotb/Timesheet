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

app.post('/get-employees', upload.single('excel'), async (req, res) => {
    try {
        const { month, year } = req.body;
        const generator = new TimesheetGenerator(req.file.path, month, year);
        await generator.processExcel();
        res.json({ employees: Object.keys(generator.employeeData) });
    } catch (e) { res.status(500).json({ error: e.message }); }
    finally { if (req.file) fs.unlinkSync(req.file.path); }
});

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
        const output = fs.createWriteStream(path.join(outputDir, zipName));
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
        if (req.file) fs.unlinkSync(req.file.path);
        if (runDir) fs.rmSync(runDir, { recursive: true, force: true });
    }
});

app.post('/validate-data', upload.single('excel'), async (req, res) => {
    try {
        const { month, year } = req.body;
        const validator = new DataValidator(req.file.path, month, year);
        await validator.process();
        const fileName = `Validation_${Date.now()}.xlsx`;
        await validator.generateReport(path.join(outputDir, fileName));
        res.json({ success: true, fileName });
    } catch (e) { res.status(500).json({ error: e.message }); }
    finally { if (req.file) fs.unlinkSync(req.file.path); }
});

app.post('/generate-project-summary', upload.single('excel'), async (req, res) => {
    try {
        const { month, year } = req.body;
        const gen = new ProjectSummaryGenerator(req.file.path, month, year);
        await gen.loadTimesheetData();
        const fileName = `Summary_${Date.now()}.xlsx`;
        await gen.generateExcelReport(path.join(outputDir, fileName));
        res.json({ success: true, fileName });
    } catch (e) { res.status(500).json({ error: e.message }); }
    finally { if (req.file) fs.unlinkSync(req.file.path); }
});

app.get(['/download-all', '/download-food-allowance'], (req, res) => {
    const filePath = path.join(outputDir, req.query.fileName);
    if (fs.existsSync(filePath)) res.download(filePath);
    else res.status(404).send('File not found');
});

setInterval(() => {
    const now = Date.now();
    fs.readdirSync(outputDir).forEach(f => {
        const p = path.join(outputDir, f);
        if (now - fs.statSync(p).mtimeMs > 1800000) fs.rmSync(p, { recursive: true, force: true });
    });
}, 600000);

app.listen(PORT, () => console.log(`Server started on ${PORT}`));
