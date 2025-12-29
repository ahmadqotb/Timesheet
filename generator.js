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
        } catch (error) {
            console.log('Logo.png not found');
        }
        return null;
    }

    formatDateLocal(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
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
            if (isNaN(date.getTime())) return;
            if (date.getMonth() + 1 !== this.month || date.getFullYear() !== this.year) return;

            if (!this.employeeData[employeeName]) this.employeeData[employeeName] = [];
            this.employeeData[employeeName].push({
                projectCode,
                projectName,
                date,
                dayOfWeek: date.getDay()
            });
        });

        this.calculateSummaries();
        return Object.keys(this.employeeData).length;
    }

    calculateSummaries() {
        const daysInMonth = new Date(this.year, this.month, 0).getDate();
        for (const employeeName in this.employeeData) {
            const entries = this.employeeData[employeeName];
            const workedDates = new Set();
            let workedFridays = 0;
            let workedSaturdays = 0;

            entries.forEach(e => {
                workedDates.add(this.formatDateLocal(e.date));
                if (e.dayOfWeek === 5) workedFridays++;
                if (e.dayOfWeek === 6) workedSaturdays++;
            });

            const hasSatFriOff = this.satFriEmployees.includes(employeeName);
            const totalWorkedDays = workedDates.size;
            
            // Basic logic for absence (customizable)
            let absence = daysInMonth - totalWorkedDays;
            this.employeeSummaries[employeeName] = {
                totalAbsentDays: Math.max(0, absence),
                totalPayrunDays: 30 - Math.max(0, absence),
                workedFridays,
                workedSaturdays,
                hasSatFriOff
            };
        }
    }

    generateHTML(employeeName, entries, summary) {
        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
        let rows = entries.map(e => `
            <tr>
                <td>${this.formatDateLocal(e.date)}</td>
                <td>${e.projectCode}</td>
                <td>${e.projectName}</td>
                <td>Worked</td>
            </tr>
        `).join('');

        return `
            <html>
                <style>
                    body { font-family: sans-serif; padding: 20px; font-size: 12px; }
                    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                    th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
                    th { background: #f4f4f4; }
                    .header { display: flex; justify-content: space-between; }
                </style>
                <body>
                    <div class="header">
                        <h1>Timesheet: ${employeeName}</h1>
                        ${this.logoBase64 ? `<img src="${this.logoBase64}" width="100">` : ''}
                    </div>
                    <p>Month: ${monthNames[this.month-1]} ${this.year}</p>
                    <table>
                        <tr><th>Absent Days</th><td>${summary.totalAbsentDays}</td></tr>
                        <tr><th>Payrun Days</th><td>${summary.totalPayrunDays}</td></tr>
                    </table>
                    <table>
                        <thead><tr><th>Date</th><th>Code</th><th>Project</th><th>Status</th></tr></thead>
                        <tbody>${rows}</tbody>
                    </table>
                </body>
            </html>`;
    }

    async generatePDFs(outputDir, progressCallback) {
        const browser = await puppeteer.launch({
            headless: 'new',
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || null,
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
        });

        try {
            const employees = Object.keys(this.employeeData);
            for (let i = 0; i < employees.length; i++) {
                const name = employees[i];
                if (progressCallback) progressCallback({ 
                    progress: Math.round(((i + 1) / employees.length) * 100), 
                    status: `Generating: ${name}` 
                });

                const page = await browser.newPage();
                const html = this.generateHTML(name, this.employeeData[name], this.employeeSummaries[name]);
                await page.setContent(html, { waitUntil: 'networkidle0' });
                await page.pdf({
                    path: path.join(outputDir, `${name}_Timesheet.pdf`),
                    format: 'A4',
                    printBackground: true
                });
                await page.close();
            }
        } finally {
            await browser.close();
        }
    }
}

module.exports = TimesheetGenerator;
