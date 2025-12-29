const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

class TimesheetGenerator {
    constructor(excelPath, month, year, satFriEmployees = []) {
        this.excelPath = excelPath;
        this.month = parseInt(month);
        this.year = parseInt(year);
        this.satFriEmployees = satFriEmployees; // Employees with Sat+Fri off
        this.employeeData = {};
        this.employeeSummaries = {};
        this.logoBase64 = this.getLogoBase64();
    }

    getLogoBase64() {
        try {
            const logoPath = path.join(__dirname, 'Logo.png');
            if (fs.existsSync(logoPath)) {
                const logoBuffer = fs.readFileSync(logoPath);
                return `data:image/png;base64,${logoBuffer.toString('base64')}`;
            }
        } catch (error) {
            console.log('Logo.png not found, continuing without logo');
        }
        return null;
    }

    // Format date as YYYY-MM-DD without timezone conversion
    formatDateLocal(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    async processExcel() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(this.excelPath);
        
        // Get first sheet regardless of name
        const worksheet = workbook.worksheets[0];

        if (!worksheet) {
            throw new Error('No worksheet found in Excel file');
        }

        // Skip header row (row 1)
        worksheet.eachRow((row, rowIndex) => {
            if (rowIndex === 1) return; // Skip header

            const projectCode = row.getCell(1).value;
            const projectName = row.getCell(2).value;
            const dateValue = row.getCell(3).value;
            const employeeName = row.getCell(4).value;
            const enteredBy = row.getCell(5).value;

            if (!employeeName || !dateValue) return;

            // Parse date
            let date;
            if (dateValue instanceof Date) {
                date = dateValue;
            } else if (typeof dateValue === 'string') {
                // Parse date string and force local timezone
                const parts = dateValue.split(/[-/]/);
                if (parts.length === 3) {
                    // Assume YYYY-MM-DD or YYYY/MM/DD format
                    date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
                } else {
                    date = new Date(dateValue);
                }
            } else if (typeof dateValue === 'number') {
                // Excel serial date
                date = new Date((dateValue - 25569) * 86400 * 1000);
            } else {
                return; // Skip invalid date
            }

            // Validate date
            if (isNaN(date.getTime())) {
                return;
            }

            // Check if date matches selected month/year
            if (date.getMonth() + 1 !== this.month || date.getFullYear() !== this.year) {
                return;
            }

            // Initialize employee data
            if (!this.employeeData[employeeName]) {
                this.employeeData[employeeName] = [];
            }

            this.employeeData[employeeName].push({
                projectCode: projectCode || '--',
                projectName: projectName || '--',
                date: date,
                dayOfWeek: date.getDay(),
                enteredBy: enteredBy || '--'
            });
        });

        // Calculate summaries for each employee
        this.calculateSummaries();

        return Object.keys(this.employeeData).length;
    }

    calculateSummaries() {
        const daysInMonth = new Date(this.year, this.month, 0).getDate();
        const totalFridaysInMonth = this.getTotalFridaysInMonth();
        const totalSaturdaysInMonth = this.getTotalSaturdaysInMonth();

        for (const employeeName in this.employeeData) {
            const entries = this.employeeData[employeeName];
            
            let workedFridays = 0;
            let workedSaturdays = 0;
            const workedDates = new Set();

            entries.forEach(entry => {
                const dateStr = this.formatDateLocal(entry.date);
                workedDates.add(dateStr);

                // Count worked Fridays (day 5)
                if (entry.dayOfWeek === 5) {
                    workedFridays++;
                }

                // Count worked Saturdays (day 6)
                if (entry.dayOfWeek === 6) {
                    workedSaturdays++;
                }
            });

            // Total worked days
            const totalWorkedDays = workedDates.size;

            // Check if this employee has Sat+Fri off
            const hasSatFriOff = this.satFriEmployees.includes(employeeName);

            let totalAbsentDays;

            if (hasSatFriOff) {
                // For employees with Sat+Fri off:
                // Absent = Days in Month - Worked Days - (Total Fri + Total Sat - Worked Fri - Worked Sat)
                totalAbsentDays = daysInMonth - totalWorkedDays - (totalFridaysInMonth + totalSaturdaysInMonth - workedFridays - workedSaturdays);
            } else {
                // For regular employees (only Friday off):
                // Absent = Days in Month - Worked Days - (Total Fridays - Worked Fridays)
                totalAbsentDays = daysInMonth - totalWorkedDays - (totalFridaysInMonth - workedFridays);
            }

            // Ensure absent days is not negative
            totalAbsentDays = Math.max(0, totalAbsentDays);

            // NEW LOGIC: Payrun = 30 - Absent Days
            const payrunDays = 30 - totalAbsentDays;

            this.employeeSummaries[employeeName] = {
                totalAbsentDays: totalAbsentDays,
                totalPayrunDays: payrunDays,
                workedFridays,
                workedSaturdays,
                hasSatFriOff
            };
        }
    }

    getTotalFridaysInMonth() {
        let count = 0;
        const daysInMonth = new Date(this.year, this.month, 0).getDate();
        
        for (let day = 1; day <= daysInMonth; day++) {
            const date = new Date(this.year, this.month - 1, day);
            if (date.getDay() === 5) { // Friday
                count++;
            }
        }
        
        return count;
    }

    getTotalSaturdaysInMonth() {
        let count = 0;
        const daysInMonth = new Date(this.year, this.month, 0).getDate();
        
        for (let day = 1; day <= daysInMonth; day++) {
            const date = new Date(this.year, this.month - 1, day);
            if (date.getDay() === 6) { // Saturday
                count++;
            }
        }
        
        return count;
    }

    generateHTML(employeeName, entries, summary) {
        const daysInMonth = new Date(this.year, this.month, 0).getDate();
        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                           'July', 'August', 'September', 'October', 'November', 'December'];
        const monthName = monthNames[this.month - 1];

        // Create full month calendar
        const allDays = [];
        for (let day = 1; day <= daysInMonth; day++) {
            const date = new Date(this.year, this.month - 1, day);
            allDays.push({
                day: day,
                date: date,
                dayOfWeek: date.getDay()
            });
        }

        // Map of worked dates to entries
        const workedDateMap = {};
        entries.forEach(entry => {
            const dateStr = this.formatDateLocal(entry.date);
            if (!workedDateMap[dateStr]) {
                workedDateMap[dateStr] = [];
            }
            workedDateMap[dateStr].push(entry);
        });

        const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

        // Generate table rows for all days
        let tableRows = '';
        allDays.forEach(dayInfo => {
            const dateStr = this.formatDateLocal(dayInfo.date);
            const dayName = dayNames[dayInfo.dayOfWeek];
            const isFriday = dayInfo.dayOfWeek === 5;
            const isSaturday = dayInfo.dayOfWeek === 6;
            const workedEntries = workedDateMap[dateStr];

            if (workedEntries) {
                // Worked this day
                workedEntries.forEach((entry, idx) => {
                    let rowClass = '';
                    if (isFriday) rowClass = 'friday';
                    else if (isSaturday) rowClass = 'saturday';
                    
                    tableRows += `
                        <tr class="${rowClass}">
                            <td>${idx === 0 ? dateStr : ''}</td>
                            <td>${idx === 0 ? dayName : ''}</td>
                            <td>${entry.projectCode}</td>
                            <td>${entry.projectName}</td>
                            <td>${isFriday ? 'Worked (Friday)' : (isSaturday ? 'Worked (Saturday)' : 'Worked')}</td>
                        </tr>`;
                });
            } else {
                // Did not work this day
                let status = '';
                let rowClass = '';

                if (isFriday) {
                    status = 'Off (Friday)';
                    rowClass = 'friday';
                } else if (isSaturday && summary.hasSatFriOff) {
                    status = 'Off (Saturday)';
                    rowClass = 'saturday';
                } else {
                    status = 'Absent';
                    rowClass = 'absence';
                }

                tableRows += `
                    <tr class="${rowClass}">
                        <td>${dateStr}</td>
                        <td>${dayName}</td>
                        <td>--</td>
                        <td>--</td>
                        <td>${status}</td>
                    </tr>`;
            }
        });

        return `
            <html>
                <head>
                    <style>
                        body { 
                            font-family: Arial, sans-serif; 
                            font-size: 9.5px;
                            margin: 8mm;
                        }
                        .header {
                            display: flex;
                            justify-content: space-between;
                            align-items: center;
                            margin-bottom: 6px;
                        }
                        .logo {
                            width: 70px;
                            height: auto;
                        }
                        h1 { 
                            text-align: center;
                            color: #333;
                            font-size: 17px;
                            margin: 0 0 4px 0;
                        }
                        h2 {
                            font-size: 13px;
                            margin: 10px 0 6px 0;
                            text-align: center;
                            color: #333;
                        }
                        table { 
                            width: 100%; 
                            border-collapse: collapse; 
                            font-size: 8.5px;
                            margin-bottom: 10px;
                        }
                        th, td { 
                            border: 1px solid #ddd; 
                            padding: 4px 5px;
                            text-align: left;
                        }
                        th { 
                            background-color: #00a65a; 
                            color: white;
                            text-align: center;
                            font-weight: bold;
                        }
                        .absence { 
                            background-color: #ffcccc;
                            color: #c62828;
                        }
                        .friday, .saturday { 
                            background-color: #ffe6e6;
                            color: #d32f2f;
                        }
                        .summary-table {
                            width: 100%;
                            margin-bottom: 10px;
                        }
                        .summary-table th {
                            width: 40%;
                            background-color: #f5f5f5;
                            color: #333;
                            text-align: right;
                            padding-right: 10px;
                        }
                        .summary-table td {
                            text-align: left;
                            padding-left: 10px;
                            font-weight: bold;
                        }
                        .signature { 
                            margin-top: 14px;
                            display: flex;
                            justify-content: space-between;
                            font-size: 8.5px;
                            padding: 0 15px;
                        }
                        .signature div {
                            text-align: center;
                        }
                        .signature-line {
                            border-top: 1px solid #333;
                            width: 160px;
                            margin-top: 25px;
                            margin-bottom: 4px;
                        }
                        .note { 
                            font-size: 7.5px;
                            font-style: italic;
                            text-align: center;
                            margin-top: 10px;
                            color: #666;
                        }
                    </style>
                </head>
                <body>
                    <div class="header">
                        <div style="flex: 1;"></div>
                        ${this.logoBase64 ? `<img src="${this.logoBase64}" class="logo" />` : ''}
                    </div>
                    <h1>Employee Timesheet Summary</h1>
                    <h2>Summary - ${monthName} ${this.year}</h2>
                    <table class="summary-table">
                        <tr>
                            <th>Employee Name</th>
                            <td>${employeeName}</td>
                        </tr>
                        <tr>
                            <th>Total Absent Days</th>
                            <td>${summary.totalAbsentDays}</td>
                        </tr>
                        <tr>
                            <th>Worked Fridays</th>
                            <td>${summary.workedFridays}</td>
                        </tr>
                        <tr>
                            <th>Total Payrun Days</th>
                            <td>${summary.totalPayrunDays}</td>
                        </tr>
                    </table>
                    
                    <h2>Daily Timesheet</h2>
                    <table>
                        <thead>
                            <tr>
                                <th>Date</th>
                                <th>Day</th>
                                <th>Project Code</th>
                                <th>Project Name</th>
                                <th>Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${tableRows}
                        </tbody>
                    </table>
                    
                    <div class="signature">
                        <div>
                            <div class="signature-line"></div>
                            <div>Approved By: Managing Director</div>
                        </div>
                        <div>
                            <div class="signature-line"></div>
                            <div>Reviewed By: TSR Operation Manager</div>
                        </div>
                    </div>
                    
                    <div class="note">
                        This document has been generated automatically from the timesheet entries system.
                    </div>
                </body>
            </html>
        `;
    }

    async generatePDFs(outputDir, progressCallback) {
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

// 1. Define possible paths where Railway/Nixpacks installs Chromium
        const possiblePaths = [
            process.env.PUPPETEER_EXECUTABLE_PATH,
            '/usr/bin/chromium',
            '/usr/bin/google-chrome',
            '/usr/bin/chromium-browser'
        ];

        // 2. Find the first path that actually exists
        const fs = require('fs');
        const executablePath = possiblePaths.find(path => path && fs.existsSync(path));

        console.log(`[PDF Generator] Using browser at: ${executablePath || 'NOT FOUND'}`);

        const browser = await puppeteer.launch({
            headless: 'new',
            executablePath: executablePath, 
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-gpu',
                '--single-process'
            ]
        });

        const employees = Object.keys(this.employeeData);
        const total = employees.length;

        for (let i = 0; i < total; i++) {
            const employeeName = employees[i];
            const entries = this.employeeData[employeeName];
            const summary = this.employeeSummaries[employeeName];

            if (progressCallback) {
                progressCallback({
                    progress: Math.round(((i + 1) / total) * 100),
                    status: `Generating PDF ${i + 1}/${total}: ${employeeName}`
                });
            }

            const html = this.generateHTML(employeeName, entries, summary);
            const page = await browser.newPage();
            await page.setContent(html);

            const fileName = `${employeeName}_Timesheet_${String(this.month).padStart(2, '0')}-${this.year}.pdf`;
            const filePath = path.join(outputDir, fileName);

            await page.pdf({
                path: filePath,
                format: 'A4',
                printBackground: true,
                margin: {
                    top: '8mm',
                    right: '9mm',
                    bottom: '8mm',
                    left: '9mm'
                }
            });

            await page.close();
        }

        await browser.close();

        return total;
    }
}

module.exports = TimesheetGenerator;
