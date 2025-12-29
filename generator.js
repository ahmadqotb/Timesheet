const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

class TimesheetGenerator {
    constructor(excelPath, month, year, satFriEmployees = []) {
        this.excelPath = excelPath;
        this.month = parseInt(month);
        this.year = parseInt(year);
        this.satFriEmployees = Array.isArray(satFriEmployees) ? satFriEmployees : [];
        this.employeeData = {};
        this.employeeSummaries = {};
        this.logoBase64 = this.getLogoBase64();
    }

    getLogoBase64() {
        try {
            const logoPath = path.join(__dirname, 'Logo.png');
            if (fs.existsSync(logoPath)) {
                return `data:image/png;base64,${fs.readFileSync(logoPath).toString('base64')}`;
            }
        } catch (e) {}
        return null;
    }

    formatDateLocal(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    async processExcel() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(this.excelPath);
        const worksheet = workbook.worksheets[0];

        worksheet.eachRow((row, rowIndex) => {
            if (rowIndex === 1) return;
            const projectCode = row.getCell(1).value || '--';
            const projectName = row.getCell(2).value || '--';
            const dateValue = row.getCell(3).value;
            const employeeName = row.getCell(4).value;

            if (!employeeName || !dateValue) return;

            let date = dateValue instanceof Date ? dateValue : new Date(dateValue);
            if (isNaN(date.getTime()) || date.getMonth() + 1 !== this.month || date.getFullYear() !== this.year) return;

            if (!this.employeeData[employeeName]) this.employeeData[employeeName] = [];
            this.employeeData[employeeName].push({ projectCode, projectName, date, dayOfWeek: date.getDay() });
        });

        this.calculateSummaries();
        return Object.keys(this.employeeData).length;
    }

    calculateSummaries() {
        const daysInMonth = new Date(this.year, this.month, 0).getDate();
        for (const employeeName in this.employeeData) {
            const entries = this.employeeData[employeeName];
            const workedDates = new Set();
            entries.forEach(e => workedDates.add(this.formatDateLocal(e.date)));
            
            const totalWorkedDays = workedDates.size;
            let absence = daysInMonth - totalWorkedDays;
            
            this.employeeSummaries[employeeName] = {
                totalAbsentDays: Math.max(0, absence),
                totalPayrunDays: 30 - Math.max(0, absence)
            };
        }
    }

    generateHTML(name, entries, summary) {
        const rows = entries.map(e => `
            <tr>
                <td>${this.formatDateLocal(e.date)}</td>
                <td>${e.projectCode}</td>
                <td>${e.projectName}</td>
                <td>Worked</td>
            </tr>`).join('');

        return `<html><style>body{font-family:sans-serif;font-size:12px;}table{width:100%;border-collapse:collapse;}th,td{border:1px solid #ccc;padding:5px;}</style>
                <body><h1>Timesheet: ${name}</h1><p>Absent: ${summary.totalAbsentDays}</p>
                <table><thead><tr><th>Date</th><th>Code</th><th>Project</th><th>Status</th></tr></thead>
                <tbody>${rows}</tbody></table></body></html>`;
    }

    async generatePDFs(outputDir, progressCallback) {
        const browser = await puppeteer.launch({
            headless: 'new',
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || '/usr/bin/chromium',
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
        });

        try {
            const employees = Object.keys(this.employeeData);
            for (let i = 0; i < employees.length; i++) {
                const name = employees[i];
                if (progressCallback) progressCallback({ progress: Math.round(((i+1)/employees.length)*100), status: `Printing: ${name}` });

                const page = await browser.newPage();
                try {
                    const html = this.generateHTML(name, this.employeeData[name], this.employeeSummaries[name]);
                    await page.setContent(html, { waitUntil: 'networkidle0' });
                    await page.pdf({ path: path.join(outputDir, `${name}_Timesheet.pdf`), format: 'A4', printBackground: true });
                } finally {
                    await page.close();
                }
            }
        } finally {
            await browser.close();
        }
    }
}

module.exports = TimesheetGenerator;
