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
    }

    async processExcel() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(this.excelPath);
        const worksheet = workbook.worksheets[0];
        worksheet.eachRow((row, rowIndex) => {
            if (rowIndex === 1) return;
            const dateValue = row.getCell(3).value;
            const employeeName = row.getCell(4).value;
            if (!employeeName || !dateValue) return;
            let date = dateValue instanceof Date ? dateValue : new Date(dateValue);
            if (date.getMonth() + 1 === this.month && date.getFullYear() === this.year) {
                if (!this.employeeData[employeeName]) this.employeeData[employeeName] = [];
                this.employeeData[employeeName].push({
                    projectCode: row.getCell(1).value || '--',
                    projectName: row.getCell(2).value || '--',
                    date
                });
            }
        });
        for (const name in this.employeeData) {
            this.employeeSummaries[name] = { workedDays: this.employeeData[name].length };
        }
        return Object.keys(this.employeeData).length;
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
                if (progressCallback) progressCallback({ progress: Math.round(((i + 1) / employees.length) * 100), status: `Printing: ${name}` });
                const page = await browser.newPage();
                const html = `<html><body><h1>${name}</h1></body></html>`;
                await page.setContent(html);
                await page.pdf({ path: path.join(outputDir, `${name}.pdf`), format: 'A4' });
                await page.close();
            }
        } finally {
            await browser.close();
        }
    }
}

module.exports = TimesheetGenerator;
