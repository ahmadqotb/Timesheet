const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

class AbsenceReportGenerator {
    constructor(timesheetPath, month, year, satFriEmployees) {
        this.timesheetPath = timesheetPath;
        this.month = parseInt(month);
        this.year = parseInt(year);
        this.satFriEmployees = satFriEmployees || [];
        
        this.employeeData = {}; // {EmployeeName: [{date, projectCode, projectName, isAnnualLeave}]}
        this.payrunSummary = {}; // Final calculations for each employee
    }

    async process() {
        await this.loadTimesheetData();
        this.calculatePayrunAndAbsence();
    }

    formatDateLocal(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    async loadTimesheetData() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(this.timesheetPath);
        const worksheet = workbook.worksheets[0];

        worksheet.eachRow((row, rowIndex) => {
            if (rowIndex === 1) return; // Skip header

            const projectCode = row.getCell(1).value;
            const projectName = row.getCell(2).value;
            const dateValue = row.getCell(3).value;
            const employeeName = row.getCell(4).value;

            if (!employeeName || !dateValue) return;

            // Parse date
            let date;
            if (dateValue instanceof Date) {
                date = dateValue;
            } else if (typeof dateValue === 'string') {
                const parts = dateValue.split(/[-/]/);
                if (parts.length === 3) {
                    date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
                } else {
                    date = new Date(dateValue);
                }
            } else if (typeof dateValue === 'number') {
                date = new Date((dateValue - 25569) * 86400 * 1000);
            } else {
                return;
            }

            if (isNaN(date.getTime())) return;

            // Filter by month/year
            if (date.getMonth() + 1 !== this.month || date.getFullYear() !== this.year) {
                return;
            }

            // Check if Annual Leave (Project Code = "A/L")
            const isAnnualLeave = projectCode && projectCode.toString().toUpperCase() === 'A/L';

            // Initialize employee data array
            if (!this.employeeData[employeeName]) {
                this.employeeData[employeeName] = [];
            }

            this.employeeData[employeeName].push({
                date: date,
                projectCode: projectCode || '',
                projectName: projectName || '',
                isAnnualLeave: isAnnualLeave
            });
        });
    }

    getDayOfWeek(date) {
        return date.getDay(); // 0 = Sunday, 5 = Friday, 6 = Saturday
    }

    calculatePayrunAndAbsence() {
        // Get days in month
        const daysInMonth = new Date(this.year, this.month, 0).getDate();

        // Count Fridays and Saturdays in month
        let totalFridays = 0;
        let totalSaturdays = 0;
        for (let day = 1; day <= daysInMonth; day++) {
            const date = new Date(this.year, this.month - 1, day);
            const dayOfWeek = this.getDayOfWeek(date);
            if (dayOfWeek === 5) totalFridays++; // Friday
            if (dayOfWeek === 6) totalSaturdays++; // Saturday
        }

        // Process each employee
        for (const employeeName in this.employeeData) {
            const entries = this.employeeData[employeeName];
            const isSatFriEmployee = this.satFriEmployees.includes(employeeName);

            // Group entries by date
            const dateMap = {};
            entries.forEach(entry => {
                const dateStr = this.formatDateLocal(entry.date);
                if (!dateMap[dateStr]) {
                    dateMap[dateStr] = [];
                }
                dateMap[dateStr].push(entry);
            });

            // Count worked days and Fridays/Saturdays
            let workedDays = 0;
            let workedFridays = 0;
            let workedSaturdays = 0;

            for (const dateStr in dateMap) {
                const date = new Date(dateStr);
                const dayOfWeek = this.getDayOfWeek(date);

                workedDays++;

                if (dayOfWeek === 5) { // Friday
                    workedFridays++;
                }
                if (dayOfWeek === 6) { // Saturday
                    workedSaturdays++;
                }
            }

            // Calculate absence
            let absence;
            if (isSatFriEmployee) {
                // Sat+Fri employees: Saturdays and Fridays are OFF days
                // Absent = Days in Month - Worked Days - (Total Fri + Total Sat - Worked Fri - Worked Sat)
                absence = daysInMonth - workedDays - (totalFridays + totalSaturdays - workedFridays - workedSaturdays);
            } else {
                // Regular employees: Only Fridays are OFF
                // Absent = Days in Month - Worked Days - (Total Fridays - Worked Fridays)
                absence = daysInMonth - workedDays - (totalFridays - workedFridays);
            }

            // Ensure absence is not negative
            absence = Math.max(0, absence);

            // NEW LOGIC: Payrun = 30 - Absent Days
            const payrunDays = 30 - absence;

            // Find absent dates
            const absentDates = [];
            for (let day = 1; day <= daysInMonth; day++) {
                const date = new Date(this.year, this.month - 1, day);
                const dateStr = this.formatDateLocal(date);
                const dayOfWeek = this.getDayOfWeek(date);

                // Skip Fridays for all employees
                if (dayOfWeek === 5) continue;

                // Skip Saturdays for Sat+Fri employees
                if (isSatFriEmployee && dayOfWeek === 6) continue;

                // Check if employee has record for this date
                if (!dateMap[dateStr]) {
                    absentDates.push(dateStr);
                }
            }

            this.payrunSummary[employeeName] = {
                workedDays: workedDays,
                workedFridays: workedFridays,
                workedSaturdays: workedSaturdays,
                payrunDays: payrunDays,
                absence: absence,
                absentDates: absentDates,
                isSatFriEmployee: isSatFriEmployee,
                dateMap: dateMap
            };
        }
    }

    async generateExcelReport(outputPath) {
        const workbook = new ExcelJS.Workbook();
        
        // Sheet 1: Payrun Summary
        const summarySheet = workbook.addWorksheet('Payrun Summary');
        this.createPayrunSummarySheet(summarySheet);

        // Sheet 2: Absence Summary
        const absenceSheet = workbook.addWorksheet('Absence Summary');
        this.createAbsenceSummarySheet(absenceSheet);

        // Sheet 3: Detailed Records
        const detailSheet = workbook.addWorksheet('Detailed Records');
        this.createDetailedRecordsSheet(detailSheet);

        await workbook.xlsx.writeFile(outputPath);
        return outputPath;
    }

    createPayrunSummarySheet(sheet) {
        // Headers
        sheet.columns = [
            { header: 'Employee Name', key: 'name', width: 35 },
            { header: 'Total Worked Days', key: 'worked', width: 18 },
            { header: 'Worked Fridays', key: 'fridays', width: 16 },
            { header: 'Payrun Days', key: 'payrun', width: 14 },
            { header: 'Absence', key: 'absence', width: 12 }
        ];

        // Style header row
        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.height = 25;
        headerRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF4472C4' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add data rows
        const sortedNames = Object.keys(this.payrunSummary).sort();

        sortedNames.forEach(employeeName => {
            const data = this.payrunSummary[employeeName];
            const row = sheet.addRow({
                name: employeeName,
                worked: data.workedDays,
                fridays: data.workedFridays,
                payrun: data.payrunDays,
                absence: data.absence
            });

            // Borders
            row.eachCell(cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            // Highlight if absent
            if (data.absence > 0) {
                row.getCell('absence').fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFEB3B' }
                };
                row.getCell('absence').font = { bold: true };
            }
        });
    }

    createAbsenceSummarySheet(sheet) {
        // Headers
        sheet.columns = [
            { header: 'Employee Name', key: 'name', width: 35 },
            { header: 'Total Worked Days', key: 'worked', width: 18 },
            { header: 'Absence', key: 'absence', width: 12 },
            { header: 'Absent Dates', key: 'dates', width: 60 }
        ];

        // Style header row
        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.height = 25;
        headerRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE74C3C' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add data rows (only employees with absence > 0)
        const sortedNames = Object.keys(this.payrunSummary)
            .filter(name => this.payrunSummary[name].absence > 0)
            .sort();

        sortedNames.forEach(employeeName => {
            const data = this.payrunSummary[employeeName];
            const row = sheet.addRow({
                name: employeeName,
                worked: data.workedDays,
                absence: data.absence,
                dates: data.absentDates.join(', ')
            });

            // Borders
            row.eachCell(cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            // Highlight absence count
            row.getCell('absence').fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFCCCC' }
            };
            row.getCell('absence').font = { bold: true };
        });
    }

    createDetailedRecordsSheet(sheet) {
        // Headers
        sheet.columns = [
            { header: 'Employee Name', key: 'name', width: 35 },
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Project Code', key: 'code', width: 15 },
            { header: 'Project Name', key: 'project', width: 40 },
            { header: 'Status', key: 'status', width: 28 }
        ];

        // Style header row
        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.height = 25;
        headerRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF2ECC71' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Get days in month
        const daysInMonth = new Date(this.year, this.month, 0).getDate();

        // Sort employees
        const sortedNames = Object.keys(this.payrunSummary).sort();

        // Add data rows
        sortedNames.forEach(employeeName => {
            const data = this.payrunSummary[employeeName];
            const isSatFriEmployee = data.isSatFriEmployee;
            const dateMap = data.dateMap;

            // For each day in month
            for (let day = 1; day <= daysInMonth; day++) {
                const date = new Date(this.year, this.month - 1, day);
                const dateStr = this.formatDateLocal(date);
                const dayOfWeek = this.getDayOfWeek(date);

                // Skip Fridays for all employees
                if (dayOfWeek === 5) continue;

                // Skip Saturdays for Sat+Fri employees ONLY if they didn't work that day
                if (isSatFriEmployee && dayOfWeek === 6 && !dateMap[dateStr]) continue;

                // Check if employee has record for this date
                if (dateMap[dateStr]) {
                    // Has record - Actual Records
                    const dayEntries = dateMap[dateStr];
                    // If multiple projects, combine them
                    const projectCodes = dayEntries.map(e => e.projectCode).filter(c => c).join(', ');
                    const projectNames = dayEntries.map(e => e.projectName).filter(n => n).join(', ');

                    const row = sheet.addRow({
                        name: employeeName,
                        date: dateStr,
                        code: projectCodes,
                        project: projectNames,
                        status: 'Actual Records'
                    });

                    // Green background for actual records
                    row.getCell('status').fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFC8E6C9' }
                    };

                    row.eachCell(cell => {
                        cell.border = {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                        };
                    });
                } else {
                    // No record - Auto Filled Possible Absent
                    const row = sheet.addRow({
                        name: employeeName,
                        date: dateStr,
                        code: '',
                        project: '',
                        status: 'Auto Filled Possible Absent'
                    });

                    // Yellow background for possible absent
                    row.getCell('status').fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFF9C4' }
                    };

                    row.eachCell(cell => {
                        cell.border = {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                        };
                    });
                }
            }
        });
    }
}

module.exports = AbsenceReportGenerator;
