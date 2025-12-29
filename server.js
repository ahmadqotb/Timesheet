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

// Create necessary directories
const uploadsDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');

if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}

if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadsDir);
    },
    filename: (req, file, cb) => {
        const uniqueName = Date.now() + '-' + file.originalname;
        cb(null, uniqueName);
    }
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        if (file.originalname.match(/\.(xlsx|xls)$/)) {
            cb(null, true);
        } else {
            cb(new Error('Only Excel files are allowed!'));
        }
    }
});

// Serve static files
app.use(express.static(__dirname));

// Main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Get employee names from Excel
app.post('/get-employees', upload.single('excel'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const { month, year } = req.body;
        const generator = new TimesheetGenerator(req.file.path, month, year);
        
        await generator.processExcel();
        const employees = Object.keys(generator.employeeData);
        
        if (req.file) fs.unlinkSync(req.file.path);
        
        res.json({ employees });
    } catch (error) {
        console.error('Error:', error);
        if (req.file) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: error.message });
    }
});

// Generate PDFs endpoint
app.post('/generate', upload.single('excel'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const { month, year } = req.body;

        if (!month || !year) {
            return res.status(400).json({ error: 'Month and year are required' });
        }

        const satFriEmployees = req.body.satFriEmployees ? JSON.parse(req.body.satFriEmployees) : [];

        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');

        const sendProgress = (data) => {
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        };

        const files = fs.readdirSync(outputDir);
        files.forEach(file => {
            const filePath = path.join(outputDir, file);
            if (!fs.lstatSync(filePath).isDirectory()) {
                fs.unlinkSync(filePath);
            }
        });

        sendProgress({ progress: 0, status: 'Processing Excel file...' });

        const generator = new TimesheetGenerator(req.file.path, month, year, satFriEmployees);
        
        sendProgress({ progress: 10, status: 'Reading Excel data...' });
        const employeeCount = await generator.processExcel();

        if (employeeCount === 0) {
            sendProgress({ error: 'No data found for the selected month/year' });
            fs.unlinkSync(req.file.path);
            return res.end();
        }

        sendProgress({ progress: 20, status: `Found ${employeeCount} employees. Generating PDFs...` });

        // The generator.js needs to be updated to use /usr/bin/google-chrome
        await generator.generatePDFs(outputDir, (progress) => {
            sendProgress(progress);
        });

        fs.unlinkSync(req.file.path);

        sendProgress({
            progress: 100,
            status: 'Complete!',
            complete: true,
            count: employeeCount
        });

        res.end();

    } catch (error) {
        console.error('Error:', error);
        res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
        res.end();
    }
});

// Generate Food Allowance Report endpoint
app.post('/generate-food-allowance', upload.fields([
    { name: 'timesheetExcel', maxCount: 1 },
    { name: 'projectPoliciesExcel', maxCount: 1 },
    { name: 'employeePoliciesExcel', maxCount: 1 }
]), async (req, res) => {
    try {
        if (!req.files || !req.files.timesheetExcel || !req.files.projectPoliciesExcel || !req.files.employeePoliciesExcel) {
            return res.status(400).json({ error: 'All three Excel files are required' });
        }

        const { month, year } = req.body;

        if (!month || !year) {
            return res.status(400).json({ error: 'Month and year are required' });
        }

        const timesheetPath = req.files.timesheetExcel[0].path;
        const projectPoliciesPath = req.files.projectPoliciesExcel[0].path;
        const employeePoliciesPath = req.files.employeePoliciesExcel[0].path;

        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');

        const sendProgress = (data) => {
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        };

        sendProgress({ progress: 0, status: 'Processing files...' });

        const calculator = new FoodAllowanceCalculator(
            timesheetPath,
            projectPoliciesPath,
            employeePoliciesPath,
            month,
            year
        );

        sendProgress({ progress: 20, status: 'Loading project policies...' });
        await calculator.loadProjectPolicies();

        sendProgress({ progress: 40, status: 'Loading employee policies...' });
        await calculator.loadEmployeePolicies();

        sendProgress({ progress: 60, status: 'Processing timesheet data...' });
        await calculator.loadTimesheetData();

        sendProgress({ progress: 80, status: 'Calculating food allowances...' });
        calculator.calculateFoodAllowance();

        sendProgress({ progress: 90, status: 'Generating Excel report...' });
        
        const monthYear = String(month).padStart(2, '0') + year;
        const outputFileName = `Food_Allowance_Report_${monthYear}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);
        
        await calculator.generateExcelReport(outputPath);

        fs.unlinkSync(timesheetPath);
        fs.unlinkSync(projectPoliciesPath);
        fs.unlinkSync(employeePoliciesPath);

        sendProgress({
            progress: 100,
            status: 'Complete!',
            complete: true,
            fileName: outputFileName
        });

        res.end();

    } catch (error) {
        console.error('Error:', error);
        
        if (req.files) {
            if (req.files.timesheetExcel) fs.unlinkSync(req.files.timesheetExcel[0].path);
            if (req.files.projectPoliciesExcel) fs.unlinkSync(req.files.projectPoliciesExcel[0].path);
            if (req.files.employeePoliciesExcel) fs.unlinkSync(req.files.employeePoliciesExcel[0].path);
        }
        
        res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
        res.end();
    }
});

// ... [Rest of your report endpoints remain the same] ...

// ============================================
// ALL IN ONE - Generate All Reports
// ============================================
app.post('/generate-all', upload.fields([
    { name: 'timesheetExcel', maxCount: 1 },
    { name: 'projectPoliciesExcel', maxCount: 1 },
    { name: 'employeePoliciesExcel', maxCount: 1 }
]), async (req, res) => {
    try {
        if (!req.files || !req.files.timesheetExcel || !req.files.projectPoliciesExcel || !req.files.employeePoliciesExcel) {
            return res.status(400).json({ error: 'All three Excel files are required' });
        }

        const { month, year, satFriEmployees } = req.body;
        const timesheetPath = req.files.timesheetExcel[0].path;
        const projectPoliciesPath = req.files.projectPoliciesExcel[0].path;
        const employeePoliciesPath = req.files.employeePoliciesExcel[0].path;
        const parsedSatFriEmployees = satFriEmployees ? JSON.parse(satFriEmployees) : [];

        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');

        const sendProgress = (data) => {
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        };

        const monthYear = String(month).padStart(2, '0') + year;

        // Implementation of PDF logic and other reports [Omitted for brevity, matches your logic]

        // IMPORTANT: Ensure TimesheetGenerator in your generate-all logic 
        // also utilizes the production chrome path inside its class.
        
        res.end();

    } catch (error) {
        // [Error handling]
        res.end();
    }
});

// [Rest of server.js Download/Listen logic]
app.listen(PORT, () => {
    console.log(`🚀 Server running on: http://localhost:${PORT}`);
});
