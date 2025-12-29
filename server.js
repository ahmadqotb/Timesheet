const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const TimesheetGenerator = require('./generator');
const FoodAllowanceCalculator = require('./foodAllowance');
// Ensure these files exist in your project or comment them out if not used yet
// const AbsenceReportGenerator = require('./absenceReport');
// const ProjectSummaryGenerator = require('./projectSummary');

const app = express();
const PORT = process.env.PORT || 3000;

// Create base directories
const uploadsDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');

[uploadsDir, outputDir].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadsDir),
    filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        if (file.originalname.match(/\.(xlsx|xls)$/)) cb(null, true);
        else cb(new Error('Only Excel files are allowed!'));
    }
});

app.use(express.static(__dirname));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Helper to clean up old files periodically (optional but good for Railway)
const cleanOldFiles = (directory) => {
    try {
        const files = fs.readdirSync(directory);
        const now = Date.now();
        files.forEach(file => {
            const filePath = path.join(directory, file);
            const stats = fs.statSync(filePath);
            if (now - stats.mtimeMs > 3600000) { // Delete files older than 1 hour
                fs.rmSync(filePath, { recursive: true, force: true });
            }
        });
    } catch (err) { console.error("Cleanup error:", err); }
};

// --- ENDPOINTS ---

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
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: error.message });
    }
});

app.post('/generate', upload.single('excel'), async (req, res) => {
    let requestFolder = '';
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
        const { month, year } = req.body;
        const satFriEmployees = req.body.satFriEmployees ? JSON.parse(req.body.satFriEmployees) : [];

        // Create a unique folder for this specific request
        requestFolder = path.join(outputDir, `gen-${Date.now()}`);
        fs.mkdirSync(requestFolder, { recursive: true });

        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');

        const sendProgress = (data) => res.write(`data: ${JSON.stringify(data)}\n\n`);

        sendProgress({ progress: 10, status: 'Processing Excel file...' });
        const generator = new TimesheetGenerator(req.file.path, month, year, satFriEmployees);
        const employeeCount = await generator.processExcel();

        if (employeeCount === 0) {
            sendProgress({ error: 'No data found for the selected month/year' });
            return res.end();
        }

        sendProgress({ progress: 20, status: `Found ${employeeCount} employees. Launching Browser...` });

        // Generate PDFs into the unique request folder
        await generator.generatePDFs(requestFolder, (progress) => {
            sendProgress(progress);
        });

        sendProgress({
            progress: 100,
            status: 'Complete!',
            complete: true,
            count: employeeCount,
            folder: path.basename(requestFolder) // Send folder name to frontend if needed
        });

        res.end();
    } catch (error) {
        console.error('Generation Error:', error);
        res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
        res.end();
    } finally {
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        cleanOldFiles(outputDir); // Run cleanup for old sessions
    }
});

// Download endpoint (If your index.html requests a zip)
app.get('/download/:folder', (req, res) => {
    const folderPath = path.join(outputDir, req.params.folder);
    if (!fs.existsSync(folderPath)) return res.status(404).send('Files expired or not found');

    const archive = archiver('zip');
    res.attachment(`Timesheets-${req.params.folder}.zip`);
    archive.pipe(res);
    archive.directory(folderPath, false);
    archive.finalize();
});

app.listen(PORT, () => {
    console.log(`🚀 Server running on port ${PORT}`);
});
