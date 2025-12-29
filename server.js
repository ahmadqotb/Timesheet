const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const TimesheetGenerator = require('./generator');
const FoodAllowanceCalculator = require('./foodAllowance');
const AbsenceReportGenerator = require('./absenceReport');
const DataValidator = require('./dataValidator');
const ProjectSummaryGenerator = require('./projectSummary');

const app = express();
const PORT = process.env.PORT || 3000;

const uploadsDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');

[uploadsDir, outputDir].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadsDir),
    filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});

const upload = multer({ storage });

app.use(express.static(__dirname));

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

// Helper for Progress Streaming
const createProgressStream = (res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    return (data) => res.write(`data: ${JSON.stringify(data)}\n\n`);
};

// 1. PDF GENERATION
app.post('/generate', upload.single('excel'), async (req, res) => {
    let subDir = '';
    try {
        const { month, year } = req.body;
        const satFriEmployees = req.body.satFriEmployees ? JSON.parse(req.body.satFriEmployees) : [];
        const sendProgress = createProgressStream(res);

        const timestamp = Date.now();
        subDir = path.join(outputDir, `run-${timestamp}`);
        fs.mkdirSync(subDir, { recursive: true });

        const generator = new TimesheetGenerator(req.file.path, month, year, satFriEmployees);
        const count = await generator.processExcel();
        
        await generator.generatePDFs(subDir, (p) => sendProgress(p));

        const zipName = `Timesheets_${timestamp}.zip`;
        const output = fs.createWriteStream(path.join(outputDir, zipName));
        const archive = archiver('zip');
        
        archive.pipe(output);
        archive.directory(subDir, false);
        await archive.finalize();

        sendProgress({ progress: 100, status: 'Complete!', complete: true, count, zipFile: zipName });
        res.end();
    } catch (error) {
        res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
        res.end();
    } finally {
        if (req.file) fs.unlinkSync(req.file.path);
        if (subDir) fs.rmSync(subDir, { recursive: true, force: true });
    }
});

// 2. DATA VALIDATOR
app.post('/validate-data', upload.single('excel'), async (req, res) => {
    try {
        const { month, year } = req.body;
        const validator = new DataValidator(req.file.path, month, year);
        await validator.process();
        const fileName = `Validation_Report_${Date.now()}.xlsx`;
        await validator.generateReport(path.join(outputDir, fileName));
        res.json({ success: true, fileName });
    } catch (error) { res.status(500).json({ error: error.message }); }
    finally { if (req.file) fs.unlinkSync(req.file.path); }
});

// 3. PROJECT SUMMARY
app.post('/generate-project-summary', upload.single('excel'), async (req, res) => {
    try {
        const { month, year } = req.body;
        const generator = new ProjectSummaryGenerator(req.file.path, month, year);
        await generator.loadTimesheetData();
        const fileName = `Project_Summary_${Date.now()}.xlsx`;
        await generator.generateExcelReport(path.join(outputDir, fileName));
        res.json({ success: true, fileName });
    } catch (error) { res.status(500).json({ error: error.message }); }
    finally { if (req.file) fs.unlinkSync(req.file.path); }
});

// DOWNLOADS
app.get('/download-all', (req, res) => {
    const filePath = path.join(outputDir, req.query.fileName);
    if (fs.existsSync(filePath)) res.download(filePath);
    else res.status(404).send('File expired');
});

app.get('/download-food-allowance', (req, res) => {
    const filePath = path.join(outputDir, req.query.fileName);
    if (fs.existsSync(filePath)) res.download(filePath);
    else res.status(404).send('File expired');
});

// CLEANUP (Every 30 mins)
setInterval(() => {
    const now = Date.now();
    fs.readdirSync(outputDir).forEach(f => {
        const p = path.join(outputDir, f);
        if (now - fs.statSync(p).mtimeMs > 1800000) fs.rmSync(p, { recursive: true, force: true });
    });
}, 600000);

app.listen(PORT, () => console.log(`Server on port ${PORT}`));
