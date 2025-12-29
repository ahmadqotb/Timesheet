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
        const TimesheetGenerator = require('./generator');
        const generator = new TimesheetGenerator(req.file.path, month, year);
        
        await generator.processExcel();
        const employees = Object.keys(generator.employeeData);
        
        // Clean up uploaded file
        fs.unlinkSync(req.file.path);
        
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
            fs.unlinkSync(path.join(outputDir, file));
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

// Generate Absence Report endpoint
app.post('/generate-absence-report', upload.single('timesheetExcel'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Timesheet Excel file is required' });
        }

        const { month, year, satFriEmployees } = req.body;

        if (!month || !year) {
            return res.status(400).json({ error: 'Month and year are required' });
        }

        const timesheetPath = req.file.path;
        const parsedSatFriEmployees = satFriEmployees ? JSON.parse(satFriEmployees) : [];

        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');

        const sendProgress = (data) => {
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        };

        sendProgress({ progress: 0, status: 'Processing timesheet...' });

        const generator = new AbsenceReportGenerator(
            timesheetPath,
            month,
            year,
            parsedSatFriEmployees
        );

        sendProgress({ progress: 20, status: 'Loading timesheet data...' });
        await generator.loadTimesheetData();

        sendProgress({ progress: 50, status: 'Calculating payrun and absences...' });
        generator.calculatePayrunAndAbsence();

        sendProgress({ progress: 80, status: 'Generating Excel report...' });
        
        const monthYear = String(month).padStart(2, '0') + year;
        const outputFileName = `Absence_Report_${monthYear}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);
        
        await generator.generateExcelReport(outputPath);

        fs.unlinkSync(timesheetPath);

        sendProgress({
            progress: 100,
            status: 'Complete!',
            complete: true,
            fileName: outputFileName
        });

        res.end();

    } catch (error) {
        console.error('Error:', error);
        if (req.file) fs.unlinkSync(req.file.path);
        res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
        res.end();
    }
});

// Data Validation endpoint
app.post('/validate-data', upload.single('timesheetExcel'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Timesheet Excel file is required' });
        }

        const { month, year } = req.body;

        if (!month || !year) {
            return res.status(400).json({ error: 'Month and year are required' });
        }

        const timesheetPath = req.file.path;

        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');

        const sendProgress = (data) => {
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        };

        sendProgress({ progress: 0, status: 'Loading timesheet data...' });

        const validator = new DataValidator(timesheetPath, month, year);

        sendProgress({ progress: 20, status: 'Checking for duplicates...' });
        await validator.process();

        sendProgress({ progress: 60, status: 'Checking for inconsistencies...' });
        
        sendProgress({ progress: 80, status: 'Generating validation report...' });
        
        const monthYear = String(month).padStart(2, '0') + year;
        const outputFileName = `Data_Validation_Report_${monthYear}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);
        
        await validator.generateExcelReport(outputPath);

        fs.unlinkSync(timesheetPath);

        sendProgress({
            progress: 100,
            status: 'Complete!',
            complete: true,
            fileName: outputFileName,
            statistics: validator.statistics
        });

        res.end();

    } catch (error) {
        console.error('Error:', error);
        if (req.file) fs.unlinkSync(req.file.path);
        res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
        res.end();
    }
});

// Generate Project Summary Report endpoint
app.post('/generate-project-summary', upload.single('timesheetExcel'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Timesheet Excel file is required' });
        }

        const { month, year } = req.body;

        if (!month || !year) {
            return res.status(400).json({ error: 'Month and year are required' });
        }

        const timesheetPath = req.file.path;

        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');

        const sendProgress = (data) => {
            res.write(`data: ${JSON.stringify(data)}\n\n`);
        };

        sendProgress({ progress: 0, status: 'Loading timesheet data...' });

        const generator = new ProjectSummaryGenerator(timesheetPath, month, year);

        sendProgress({ progress: 30, status: 'Processing employee project data...' });
        const employeeCount = await generator.loadTimesheetData();

        if (employeeCount === 0) {
            sendProgress({ error: 'No data found for the selected month/year' });
            fs.unlinkSync(timesheetPath);
            return res.end();
        }

        sendProgress({ progress: 60, status: 'Calculating project percentages...' });

        sendProgress({ progress: 80, status: 'Generating Excel report with 3 sheets...' });
        
        const monthYear = String(month).padStart(2, '0') + year;
        const outputFileName = `Project_Summary_${monthYear}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);
        
        await generator.generateExcelReport(outputPath);

        fs.unlinkSync(timesheetPath);

        const statistics = generator.getStatistics();

        sendProgress({
            progress: 100,
            status: 'Complete!',
            complete: true,
            fileName: outputFileName,
            statistics: statistics
        });

        res.end();

    } catch (error) {
        console.error('Error:', error);
        if (req.file) fs.unlinkSync(req.file.path);
        res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
        res.end();
    }
});

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

        if (!month || !year) {
            return res.status(400).json({ error: 'Month and year are required' });
        }

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

        // Clear output directory
        const existingFiles = fs.readdirSync(outputDir);
        existingFiles.forEach(file => {
            const filePath = path.join(outputDir, file);
            if (fs.lstatSync(filePath).isDirectory()) {
                fs.rmSync(filePath, { recursive: true });
            } else {
                fs.unlinkSync(filePath);
            }
        });

        const monthYear = String(month).padStart(2, '0') + year;

        // ========== 1. ABSENCE REPORT (0-15%) ==========
        sendProgress({ progress: 0, status: '📋 Generating Absence Report...' });
        
        const absenceGenerator = new AbsenceReportGenerator(timesheetPath, month, year, parsedSatFriEmployees);
        await absenceGenerator.loadTimesheetData();
        absenceGenerator.calculatePayrunAndAbsence();
        await absenceGenerator.generateExcelReport(path.join(outputDir, `Absence_Report_${monthYear}.xlsx`));
        
        sendProgress({ progress: 15, status: '✅ Absence Report complete!' });

        // ========== 2. DATA VALIDATION (15-30%) ==========
        sendProgress({ progress: 16, status: '🔍 Generating Data Validation Report...' });
        
        const validator = new DataValidator(timesheetPath, month, year);
        await validator.process();
        await validator.generateExcelReport(path.join(outputDir, `Data_Validation_Report_${monthYear}.xlsx`));
        
        sendProgress({ progress: 30, status: '✅ Data Validation complete!' });

        // ========== 3. PROJECT SUMMARY (30-45%) ==========
        sendProgress({ progress: 31, status: '📈 Generating Project Summary Report...' });
        
        const projectSummary = new ProjectSummaryGenerator(timesheetPath, month, year);
        await projectSummary.loadTimesheetData();
        await projectSummary.generateExcelReport(path.join(outputDir, `Project_Summary_${monthYear}.xlsx`));
        
        sendProgress({ progress: 45, status: '✅ Project Summary complete!' });

        // ========== 4. FOOD ALLOWANCE (45-60%) ==========
        sendProgress({ progress: 46, status: '🍽️ Generating Food Allowance Report...' });
        
        const foodCalculator = new FoodAllowanceCalculator(
            timesheetPath,
            projectPoliciesPath,
            employeePoliciesPath,
            month,
            year
        );
        await foodCalculator.loadProjectPolicies();
        await foodCalculator.loadEmployeePolicies();
        await foodCalculator.loadTimesheetData();
        foodCalculator.calculateFoodAllowance();
        await foodCalculator.generateExcelReport(path.join(outputDir, `Food_Allowance_Report_${monthYear}.xlsx`));
        
        sendProgress({ progress: 60, status: '✅ Food Allowance complete!' });

        // ========== 5. TIMESHEET PDFs (60-95%) ==========
        sendProgress({ progress: 61, status: '📊 Generating Timesheet PDFs...' });
        
        const pdfDir = path.join(outputDir, 'Timesheet_PDFs');
        if (!fs.existsSync(pdfDir)) {
            fs.mkdirSync(pdfDir, { recursive: true });
        }

        const pdfGenerator = new TimesheetGenerator(timesheetPath, month, year, parsedSatFriEmployees);
        const employeeCount = await pdfGenerator.processExcel();

        await pdfGenerator.generatePDFs(pdfDir, (progress) => {
            const scaledProgress = 61 + Math.round((progress.progress / 100) * 34);
            sendProgress({ progress: scaledProgress, status: progress.status });
        });

        sendProgress({ progress: 95, status: '✅ Timesheet PDFs complete!' });

        // ========== 6. CREATE ZIP (95-100%) ==========
        sendProgress({ progress: 96, status: '📦 Creating ZIP file...' });

        const zipFileName = `All_Reports_${monthYear}.zip`;
        const zipPath = path.join(outputDir, zipFileName);

        await new Promise((resolve, reject) => {
            const output = fs.createWriteStream(zipPath);
            const archive = archiver('zip', { zlib: { level: 9 } });

            output.on('close', resolve);
            archive.on('error', reject);

            archive.pipe(output);

            // Add Excel files
            archive.file(path.join(outputDir, `Absence_Report_${monthYear}.xlsx`), { name: `Absence_Report_${monthYear}.xlsx` });
            archive.file(path.join(outputDir, `Data_Validation_Report_${monthYear}.xlsx`), { name: `Data_Validation_Report_${monthYear}.xlsx` });
            archive.file(path.join(outputDir, `Project_Summary_${monthYear}.xlsx`), { name: `Project_Summary_${monthYear}.xlsx` });
            archive.file(path.join(outputDir, `Food_Allowance_Report_${monthYear}.xlsx`), { name: `Food_Allowance_Report_${monthYear}.xlsx` });

            // Add PDF folder
            archive.directory(pdfDir, 'Timesheet_PDFs');

            archive.finalize();
        });

        // Clean up uploaded files
        fs.unlinkSync(timesheetPath);
        fs.unlinkSync(projectPoliciesPath);
        fs.unlinkSync(employeePoliciesPath);

        sendProgress({
            progress: 100,
            status: '🎉 All reports generated successfully!',
            complete: true,
            fileName: zipFileName,
            employeeCount: employeeCount
        });

        res.end();

    } catch (error) {
        console.error('Error:', error);
        
        if (req.files) {
            try { if (req.files.timesheetExcel) fs.unlinkSync(req.files.timesheetExcel[0].path); } catch(e) {}
            try { if (req.files.projectPoliciesExcel) fs.unlinkSync(req.files.projectPoliciesExcel[0].path); } catch(e) {}
            try { if (req.files.employeePoliciesExcel) fs.unlinkSync(req.files.employeePoliciesExcel[0].path); } catch(e) {}
        }
        
        res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
        res.end();
    }
});

// Download all PDFs as ZIP
app.get('/download', (req, res) => {
    const zipFileName = `Employee_Timesheets_${Date.now()}.zip`;
    
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${zipFileName}"`);

    const archive = archiver('zip', { zlib: { level: 9 } });

    archive.on('error', (err) => {
        console.error('Archive error:', err);
        res.status(500).send('Error creating archive');
    });

    archive.pipe(res);

    const files = fs.readdirSync(outputDir);
    files.forEach(file => {
        if (file.endsWith('.pdf')) {
            const filePath = path.join(outputDir, file);
            archive.file(filePath, { name: file });
        }
    });

    archive.finalize();
});

// Download Excel/ZIP files
app.get('/download-food-allowance', (req, res) => {
    const { fileName } = req.query;
    
    if (!fileName) {
        return res.status(400).send('File name is required');
    }

    const filePath = path.join(outputDir, fileName);
    
    if (!fs.existsSync(filePath)) {
        return res.status(404).send('File not found');
    }

    res.download(filePath, fileName, (err) => {
        if (err) {
            console.error('Download error:', err);
            res.status(500).send('Error downloading file');
        }
    });
});

// Start server
app.listen(PORT, () => {
    console.log('\n╔═══════════════════════════════════════════════════════╗');
    console.log('║   Employee Timesheet Generator - Server Running       ║');
    console.log('╚═══════════════════════════════════════════════════════╝\n');
    console.log(`🚀 Server is running on: http://localhost:${PORT}`);
    console.log(`📁 Upload directory: ${uploadsDir}`);
    console.log(`📄 Output directory: ${outputDir}`);
    console.log('\n✨ Open your browser and navigate to the URL above\n');
    console.log('Press Ctrl+C to stop the server\n');
});
