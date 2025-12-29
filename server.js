const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const TimesheetGenerator = require('./generator');
const FoodAllowanceCalculator = require('./foodAllowance');
// Ensure these files exist or keep them commented out
// const AbsenceReportGenerator = require('./absenceReport'); 
// const ProjectSummaryGenerator = require('./projectSummary');

const app = express();
const PORT = process.env.PORT || 3000;

// Create necessary directories
const uploadsDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');

[uploadsDir, outputDir].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadsDir),
    filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        if (file.originalname.match(/\.(xlsx|xls)$/)) cb(null, true);
        else cb(new Error('Only Excel files are allowed!'));
    }
});

app.use(express.static(__dirname));

// Main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Endpoint to get employee names for the checklist
app.post('/get-employees', upload.single('excel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
        const { month, year } = req.body;
        const generator = new TimesheetGenerator(req.file.path, month, year);
        await generator.processExcel();
        const employees = Object.keys(generator.employeeData);
        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.json({ employees });
    } catch (error) {
        console.error('Error fetching employees:', error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: error.message });
    }
});

// Generate PDFs endpoint
app.post('/generate', upload.single('excel'), async (req, res) => {
    let requestSubDir = '';
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
        
        const { month, year } = req.body;
        const satFriEmployees = req.body.satFriEmployees ? JSON.parse(req.body.satFriEmployees) : [];

        // Unique folder per request prevents users from overwriting each other's files
        const timestamp = Date.now();
        requestSubDir = path.join(outputDir, `run-${timestamp}`);
        fs.mkdirSync(requestSubDir, { recursive: true });

        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');

        const sendProgress = (data) => res.write(`data: ${JSON.stringify(data)}\n\n`);

        sendProgress({ progress: 5, status: 'Processing Excel file...' });

        const generator = new TimesheetGenerator(req.file.path, month, year, satFriEmployees);
        const employeeCount = await generator.processExcel();

        if (employeeCount === 0) {
            sendProgress({ error: 'No data found for the selected month/year' });
            return res.end();
        }

        sendProgress({ progress: 15, status: `Found ${employeeCount} employees. Generating PDFs...` });

        // Generate individual PDFs into the sub-directory
        await generator.generatePDFs(requestSubDir, (progress) => {
            sendProgress(progress);
        });

        // Create ZIP file for the "Download All" feature
        const zipFileName = `Timesheets_${String(month).padStart(2, '0')}-${year}_${timestamp}.zip`;
        const zipPath = path.join(outputDir, zipFileName);
        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip', { zlib: { level: 9 } });

        await new Promise((resolve, reject) => {
            output.on('close', resolve);
            archive.on('error', reject);
            archive.pipe(output);
            archive.directory(requestSubDir, false);
            archive.finalize();
        });

        // Cleanup: remove the individual PDFs to save disk space, keep only the ZIP
        fs.rmSync(requestSubDir, { recursive: true, force: true });
        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);

        sendProgress({
            progress: 100,
            status: 'Complete!',
            complete: true,
            count: employeeCount,
            zipFile: zipFileName // index.html uses this to create the download link
        });

        res.end();
    } catch (error) {
        console.error('Generation Crash:', error);
        res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
        res.end();
    }
});

// Generic download endpoint
app.get('/download-all', (req, res) => {
    const fileName = req.query.fileName;
    if (!fileName) return res.status(400).send('File name required');
    
    const filePath = path.join(outputDir, fileName);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).send('File not found or expired');
    }
});

// Food Allowance Download (matches index.html href)
app.get('/download-food-allowance', (req, res) => {
    const fileName = req.query.fileName;
    const filePath = path.join(outputDir, fileName);
    if (fs.existsSync(filePath)) res.download(filePath);
    else res.status(404).send('Report not found');
});

// Auto-cleanup: Deletes any output files older than 30 minutes
setInterval(() => {
    const now = Date.now();
    try {
        fs.readdirSync(outputDir).forEach(file => {
            const filePath = path.join(outputDir, file);
            const stats = fs.statSync(filePath);
            if (now - stats.mtimeMs > 1800000) { 
                fs.rmSync(filePath, { recursive: true, force: true });
                console.log(`Cleaned up old file: ${file}`);
            }
        });
    } catch (err) { console.error("Cleanup error:", err); }
}, 600000); // Runs every 10 minutes

app.listen(PORT, () => {
    console.log(`🚀 Server running at http://localhost:${PORT}`);
});
