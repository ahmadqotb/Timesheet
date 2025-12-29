const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

class DataValidator {
    constructor(timesheetPath, month, year) {
        this.timesheetPath = timesheetPath;
        this.month = parseInt(month);
        this.year = parseInt(year);
        
        this.allRecords = []; // All records with validation status
        this.duplicates = []; // Duplicate entries
        this.inconsistencies = []; // Inconsistent project codes
        this.cleanRecords = []; // Clean validated records
        this.statistics = {};
    }

    async process() {
        await this.loadTimesheetData();
        this.checkDuplicates();
        this.checkInconsistencies();
        this.generateCleanData();
        this.calculateStatistics();
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

            const projectCode = row.getCell(1).value || '';
            const projectName = row.getCell(2).value || '';
            const dateValue = row.getCell(3).value;
            const employeeName = row.getCell(4).value || '';
            const enteredBy = row.getCell(5).value || '';

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

            this.allRecords.push({
                rowNumber: rowIndex,
                projectCode: projectCode.toString().trim(),
                projectName: projectName.toString().trim(),
                date: date,
                dateStr: this.formatDateLocal(date),
                employeeName: employeeName.toString().trim(),
                enteredBy: enteredBy.toString().trim(),
                status: 'Clean', // Will be updated during validation
                statusColor: 'green' // green, yellow, red
            });
        });
    }

    checkDuplicates() {
        const seen = new Map(); // key: "Employee+Date", value: first row index
        const duplicateMap = new Map(); // key: "Employee+Date", value: array of row indices

        this.allRecords.forEach((record, index) => {
            const key = `${record.employeeName}|${record.dateStr}`;

            if (seen.has(key)) {
                // This is a duplicate
                record.status = 'Duplicate';
                record.statusColor = 'yellow';

                // Add to duplicates tracking
                if (!duplicateMap.has(key)) {
                    duplicateMap.set(key, [seen.get(key)]);
                }
                duplicateMap.get(key).push(index);
            } else {
                // First occurrence
                seen.set(key, index);
            }
        });

        // Build duplicates report
        duplicateMap.forEach((indices, key) => {
            const [employeeName, dateStr] = key.split('|');
            const records = indices.map(i => this.allRecords[i]);
            const projectCodes = records.map(r => r.projectCode).join(', ');
            const rowNumbers = records.map(r => r.rowNumber).join(', ');

            this.duplicates.push({
                employeeName: employeeName,
                date: dateStr,
                occurrences: indices.length,
                rowNumbers: rowNumbers,
                projects: projectCodes
            });
        });
    }

    normalizeProjectName(name) {
        return name.toLowerCase().trim().replace(/\s+/g, ' ');
    }

    checkInconsistencies() {
        // Group by project code
        const projectCodeMap = new Map(); // key: projectCode, value: array of records

        this.allRecords.forEach(record => {
            if (!projectCodeMap.has(record.projectCode)) {
                projectCodeMap.set(record.projectCode, []);
            }
            projectCodeMap.get(record.projectCode).push(record);
        });

        // Check each project code for inconsistencies
        projectCodeMap.forEach((records, projectCode) => {
            // Get unique normalized project names
            const uniqueNames = new Set();
            records.forEach(record => {
                const normalized = this.normalizeProjectName(record.projectName);
                uniqueNames.add(normalized);
            });

            // If more than one unique name, it's inconsistent
            if (uniqueNames.size > 1) {
                // Mark all records with this project code as inconsistent
                records.forEach(record => {
                    // Only mark as inconsistent if not already marked as duplicate
                    if (record.status === 'Clean') {
                        record.status = 'Inconsistent';
                        record.statusColor = 'red';
                    }
                });

                // Get all original (non-normalized) names for reporting
                const originalNames = [...new Set(records.map(r => r.projectName))];
                const rowNumbers = records.map(r => r.rowNumber).join(', ');

                this.inconsistencies.push({
                    projectCode: projectCode,
                    projectNames: originalNames.join(', '),
                    count: originalNames.length,
                    rowNumbers: rowNumbers
                });
            }
        });
    }

    generateCleanData() {
        // Clean data = first occurrence of each Employee+Date
        // Keep inconsistencies (only remove duplicates)
        const seen = new Set();

        this.allRecords.forEach(record => {
            const key = `${record.employeeName}|${record.dateStr}`;

            // Include if: first occurrence (regardless of inconsistencies)
            if (!seen.has(key)) {
                seen.add(key);
                this.cleanRecords.push(record);
            }
        });
    }

    calculateStatistics() {
        const total = this.allRecords.length;
        const clean = this.cleanRecords.length;
        const duplicates = this.allRecords.filter(r => r.status === 'Duplicate').length;
        const inconsistencies = this.allRecords.filter(r => r.status === 'Inconsistent').length;
        const cleanWithInconsistencies = this.cleanRecords.filter(r => r.status === 'Inconsistent').length;

        this.statistics = {
            totalRecords: total,
            cleanRecords: clean,
            duplicatesFound: duplicates,
            inconsistenciesFound: inconsistencies,
            cleanWithInconsistencies: cleanWithInconsistencies,
            validationRate: total > 0 ? ((clean / total) * 100).toFixed(1) : 0,
            duplicateRate: total > 0 ? ((duplicates / total) * 100).toFixed(1) : 0,
            inconsistencyRate: total > 0 ? ((inconsistencies / total) * 100).toFixed(1) : 0
        };
    }

    async generateExcelReport(outputPath) {
        const workbook = new ExcelJS.Workbook();

        // Sheet 1: All Data with Color Coding
        const allDataSheet = workbook.addWorksheet('All Data');
        this.createAllDataSheet(allDataSheet);

        // Sheet 2: Clean Validated Data
        const cleanSheet = workbook.addWorksheet('Clean Validated Data');
        this.createCleanDataSheet(cleanSheet);

        // Sheet 3: Duplicates Report
        const duplicatesSheet = workbook.addWorksheet('Duplicates Report');
        this.createDuplicatesSheet(duplicatesSheet);

        // Sheet 4: Inconsistencies Report
        const inconsistenciesSheet = workbook.addWorksheet('Inconsistencies Report');
        this.createInconsistenciesSheet(inconsistenciesSheet);

        // Sheet 5: Summary Report
        const summarySheet = workbook.addWorksheet('Summary');
        this.createSummarySheet(summarySheet);

        await workbook.xlsx.writeFile(outputPath);
        return outputPath;
    }

    createAllDataSheet(sheet) {
        // Headers
        sheet.columns = [
            { header: 'Project Code', key: 'code', width: 15 },
            { header: 'Project Name', key: 'name', width: 40 },
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Employee Name', key: 'employee', width: 35 },
            { header: 'Entered By', key: 'enteredBy', width: 20 },
            { header: 'Status', key: 'status', width: 15 }
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

        // Add data rows with color coding
        this.allRecords.forEach(record => {
            const row = sheet.addRow({
                code: record.projectCode,
                name: record.projectName,
                date: record.dateStr,
                employee: record.employeeName,
                enteredBy: record.enteredBy,
                status: record.status === 'Clean' ? '✅ Clean' : 
                        record.status === 'Duplicate' ? '🟡 Duplicate' : 
                        '🔴 Inconsistent'
            });

            // Color code based on status
            let fillColor;
            if (record.statusColor === 'green') {
                fillColor = 'FFC8E6C9'; // Light green
            } else if (record.statusColor === 'yellow') {
                fillColor = 'FFFFF9C4'; // Light yellow
            } else {
                fillColor = 'FFFFCCCC'; // Light red
            }

            row.eachCell(cell => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: fillColor }
                };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            // Make status cell bold
            row.getCell('status').font = { bold: true };
        });
    }

    createCleanDataSheet(sheet) {
        // Headers
        sheet.columns = [
            { header: 'Project Code', key: 'code', width: 15 },
            { header: 'Project Name', key: 'name', width: 40 },
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Employee Name', key: 'employee', width: 35 },
            { header: 'Entered By', key: 'enteredBy', width: 20 }
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
                fgColor: { argb: 'FF4CAF50' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add clean data rows
        this.cleanRecords.forEach(record => {
            const row = sheet.addRow({
                code: record.projectCode,
                name: record.projectName,
                date: record.dateStr,
                employee: record.employeeName,
                enteredBy: record.enteredBy
            });

            row.eachCell(cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });
    }

    createDuplicatesSheet(sheet) {
        // Headers
        sheet.columns = [
            { header: 'Employee Name', key: 'employee', width: 35 },
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Occurrences', key: 'occurrences', width: 15 },
            { header: 'Row Numbers', key: 'rows', width: 20 },
            { header: 'Projects', key: 'projects', width: 40 }
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
                fgColor: { argb: 'FFFFC107' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add duplicate data rows
        this.duplicates.forEach(dup => {
            const row = sheet.addRow({
                employee: dup.employeeName,
                date: dup.date,
                occurrences: dup.occurrences,
                rows: dup.rowNumbers,
                projects: dup.projects
            });

            row.eachCell(cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            // Highlight occurrences
            row.getCell('occurrences').fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFF9C4' }
            };
            row.getCell('occurrences').font = { bold: true };
        });
    }

    createInconsistenciesSheet(sheet) {
        // Headers
        sheet.columns = [
            { header: 'Project Code', key: 'code', width: 15 },
            { header: 'Project Names Found', key: 'names', width: 60 },
            { header: 'Count', key: 'count', width: 10 },
            { header: 'Row Numbers', key: 'rows', width: 30 }
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
                fgColor: { argb: 'FFF44336' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add inconsistency data rows
        this.inconsistencies.forEach(inc => {
            const row = sheet.addRow({
                code: inc.projectCode,
                names: inc.projectNames,
                count: inc.count,
                rows: inc.rowNumbers
            });

            row.eachCell(cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            // Highlight count
            row.getCell('count').fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFCCCC' }
            };
            row.getCell('count').font = { bold: true };
        });
    }

    createSummarySheet(sheet) {
        // Title
        sheet.mergeCells('A1:B1');
        const titleCell = sheet.getCell('A1');
        titleCell.value = 'DATA VALIDATION SUMMARY';
        titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
        titleCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        };
        sheet.getRow(1).height = 30;

        // Month/Year
        sheet.mergeCells('A2:B2');
        const monthCell = sheet.getCell('A2');
        monthCell.value = `Month: ${this.getMonthName(this.month)} ${this.year}`;
        monthCell.font = { bold: true };
        monthCell.alignment = { horizontal: 'center' };
        sheet.getRow(2).height = 20;

        // Empty row
        sheet.getRow(3).height = 10;

        // Statistics
        const stats = [
            ['Total Records:', this.statistics.totalRecords],
            ['Clean Records (No Duplicates):', this.statistics.cleanRecords],
            ['  - Fully Clean:', this.statistics.cleanRecords - this.statistics.cleanWithInconsistencies],
            ['  - With Inconsistencies:', this.statistics.cleanWithInconsistencies],
            ['Duplicates Removed:', this.statistics.duplicatesFound],
            ['Inconsistencies Found:', this.statistics.inconsistenciesFound],
            ['', ''],
            ['Validation Rate:', `${this.statistics.validationRate}%`],
            ['Duplicate Rate:', `${this.statistics.duplicateRate}%`],
            ['Inconsistency Rate:', `${this.statistics.inconsistencyRate}%`]
        ];

        let currentRow = 4;
        stats.forEach(([label, value]) => {
            if (label === '') {
                currentRow++;
                return;
            }

            const labelCell = sheet.getCell(`A${currentRow}`);
            const valueCell = sheet.getCell(`B${currentRow}`);

            labelCell.value = label;
            labelCell.font = { bold: true };
            labelCell.alignment = { horizontal: 'right' };

            valueCell.value = value;
            valueCell.alignment = { horizontal: 'left' };

            if (label.includes('Rate:')) {
                valueCell.font = { bold: true, color: { argb: 'FF2E7D32' } };
            }

            currentRow++;
        });

        // Column widths
        sheet.getColumn('A').width = 30;
        sheet.getColumn('B').width = 20;

        // Top issues section
        currentRow += 2;
        sheet.mergeCells(`A${currentRow}:B${currentRow}`);
        const topIssuesCell = sheet.getCell(`A${currentRow}`);
        topIssuesCell.value = 'TOP ISSUES';
        topIssuesCell.font = { bold: true, size: 14 };
        topIssuesCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };
        currentRow++;

        // Find top duplicate employees
        const employeeDupCounts = new Map();
        this.duplicates.forEach(dup => {
            employeeDupCounts.set(dup.employeeName, 
                (employeeDupCounts.get(dup.employeeName) || 0) + (dup.occurrences - 1));
        });

        const topDuplicates = [...employeeDupCounts.entries()]
            .sort((a, b) => b[1] - a[1])
            .slice(0, 3);

        topDuplicates.forEach(([employee, count]) => {
            sheet.mergeCells(`A${currentRow}:B${currentRow}`);
            const cell = sheet.getCell(`A${currentRow}`);
            cell.value = `• Employee "${employee}" has ${count} duplicate day(s)`;
            currentRow++;
        });

        // Top inconsistent codes
        const topInconsistencies = this.inconsistencies
            .sort((a, b) => b.count - a.count)
            .slice(0, 3);

        topInconsistencies.forEach(inc => {
            sheet.mergeCells(`A${currentRow}:B${currentRow}`);
            const cell = sheet.getCell(`A${currentRow}`);
            cell.value = `• Project Code "${inc.projectCode}" has ${inc.count} different names`;
            currentRow++;
        });
    }

    getMonthName(month) {
        const months = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
        return months[month - 1];
    }
}

module.exports = DataValidator;
