const ExcelJS = require('exceljs');

class ProjectSummaryGenerator {
    constructor(excelPath, month, year) {
        this.excelPath = excelPath;
        this.month = parseInt(month);
        this.year = parseInt(year);
        this.employeeProjects = {}; // { employeeName: { projectName: daysCount } }
        this.allProjects = new Set();
        this.employeeWorkedDays = {}; // { employeeName: totalWorkedDays }
    }

    async loadTimesheetData() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(this.excelPath);
        
        const worksheet = workbook.worksheets[0];
        if (!worksheet) {
            throw new Error('No worksheet found in Excel file');
        }

        // Process each row
        worksheet.eachRow((row, rowIndex) => {
            if (rowIndex === 1) return; // Skip header

            const projectName = row.getCell(2).value; // Column B - Project Name
            const dateValue = row.getCell(3).value;   // Column C - Date
            const employeeName = row.getCell(4).value; // Column D - Employee Name

            if (!employeeName || !dateValue || !projectName) return;

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

            // Validate date
            if (isNaN(date.getTime())) return;

            // Check if date matches selected month/year
            if (date.getMonth() + 1 !== this.month || date.getFullYear() !== this.year) {
                return;
            }

            // Trim and normalize project name
            const normalizedProject = String(projectName).trim();
            
            // Initialize employee data if needed
            if (!this.employeeProjects[employeeName]) {
                this.employeeProjects[employeeName] = {};
                this.employeeWorkedDays[employeeName] = new Set();
            }

            // Track unique work days per employee
            const dateKey = `${date.getFullYear()}-${date.getMonth()}-${date.getDate()}`;
            this.employeeWorkedDays[employeeName].add(dateKey);

            // Count days per project
            if (!this.employeeProjects[employeeName][normalizedProject]) {
                this.employeeProjects[employeeName][normalizedProject] = new Set();
            }
            this.employeeProjects[employeeName][normalizedProject].add(dateKey);

            // Track all unique projects
            this.allProjects.add(normalizedProject);
        });

        // Convert Sets to counts
        for (const employee in this.employeeProjects) {
            this.employeeWorkedDays[employee] = this.employeeWorkedDays[employee].size;
            for (const project in this.employeeProjects[employee]) {
                this.employeeProjects[employee][project] = this.employeeProjects[employee][project].size;
            }
        }

        return Object.keys(this.employeeProjects).length;
    }

    calculatePercentages() {
        const results = {};

        for (const employee in this.employeeProjects) {
            const totalDays = this.employeeWorkedDays[employee];
            results[employee] = {
                projects: {},
                totalDays: totalDays,
                totalPercentage: 0
            };

            for (const project of this.allProjects) {
                const projectDays = this.employeeProjects[employee][project] || 0;
                const percentage = totalDays > 0 ? (projectDays / totalDays) * 100 : 0;
                results[employee].projects[project] = {
                    days: projectDays,
                    percentage: percentage
                };
                results[employee].totalPercentage += percentage;
            }
        }

        return results;
    }

    async generateExcelReport(outputPath) {
        const workbook = new ExcelJS.Workbook();
        const percentages = this.calculatePercentages();
        
        // Sort projects alphabetically
        const sortedProjects = Array.from(this.allProjects).sort();
        
        // Sort employees alphabetically
        const sortedEmployees = Object.keys(percentages).sort();

        // ============ SHEET 1: Highlighted ============
        const sheet1 = workbook.addWorksheet('Highlighted');
        this.createSheet(sheet1, sortedEmployees, sortedProjects, percentages, 'highlighted');

        // ============ SHEET 2: With Unassigned ============
        const sheet2 = workbook.addWorksheet('With Unassigned');
        this.createSheet(sheet2, sortedEmployees, sortedProjects, percentages, 'unassigned');

        // ============ SHEET 3: Raw ============
        const sheet3 = workbook.addWorksheet('Raw');
        this.createSheet(sheet3, sortedEmployees, sortedProjects, percentages, 'raw');

        await workbook.xlsx.writeFile(outputPath);
    }

    createSheet(sheet, employees, projects, percentages, mode) {
        // Header style
        const headerStyle = {
            font: { bold: true, color: { argb: 'FFFFFFFF' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };

        // Cell style
        const cellStyle = {
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };

        // Red highlight style for totals != 100%
        const redStyle = {
            font: { bold: true, color: { argb: 'FFFF0000' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } },
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };

        // Green style for 100%
        const greenStyle = {
            font: { bold: true, color: { argb: 'FF006600' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } },
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };

        // Build headers
        const headers = ['Employee Name', ...projects];
        if (mode === 'unassigned') {
            headers.push('Unassigned');
        }
        headers.push('Total');

        // Add header row
        const headerRow = sheet.addRow(headers);
        headerRow.eachCell((cell) => {
            Object.assign(cell, headerStyle);
            cell.font = headerStyle.font;
            cell.fill = headerStyle.fill;
            cell.alignment = headerStyle.alignment;
            cell.border = headerStyle.border;
        });

        // Set column widths
        sheet.getColumn(1).width = 30; // Employee name
        for (let i = 2; i <= headers.length; i++) {
            sheet.getColumn(i).width = 15;
        }

        // Add data rows
        for (const employee of employees) {
            const empData = percentages[employee];
            const rowData = [employee];

            let rowTotal = 0;
            for (const project of projects) {
                const pct = empData.projects[project]?.percentage || 0;
                rowData.push(pct > 0 ? `${pct.toFixed(1)}%` : '0%');
                rowTotal += pct;
            }

            if (mode === 'unassigned') {
                // Calculate unassigned percentage
                const unassigned = Math.max(0, 100 - rowTotal);
                rowData.push(unassigned > 0 ? `${unassigned.toFixed(1)}%` : '0%');
                rowTotal = 100; // With unassigned, total is always 100%
            }

            // Add total column
            rowData.push(`${rowTotal.toFixed(1)}%`);

            const dataRow = sheet.addRow(rowData);
            
            // Apply styles
            dataRow.eachCell((cell, colNumber) => {
                cell.alignment = cellStyle.alignment;
                cell.border = cellStyle.border;

                // First column (employee name) left-align
                if (colNumber === 1) {
                    cell.alignment = { horizontal: 'left', vertical: 'middle' };
                }

                // Last column (Total) - apply conditional formatting
                if (colNumber === rowData.length) {
                    if (mode === 'highlighted') {
                        const totalValue = parseFloat(rowTotal.toFixed(1));
                        if (Math.abs(totalValue - 100) > 0.1) {
                            // Not 100% - red highlight
                            cell.font = redStyle.font;
                            cell.fill = redStyle.fill;
                        } else {
                            // 100% - green highlight
                            cell.font = greenStyle.font;
                            cell.fill = greenStyle.fill;
                        }
                    } else if (mode === 'unassigned') {
                        // Always green since it's always 100%
                        cell.font = greenStyle.font;
                        cell.fill = greenStyle.fill;
                    }
                    // Raw mode - no special formatting for total
                }
            });
        }

        // Freeze header row
        sheet.views = [{ state: 'frozen', ySplit: 1 }];

        // Auto-filter
        sheet.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: employees.length + 1, column: headers.length }
        };
    }

    getStatistics() {
        return {
            totalEmployees: Object.keys(this.employeeProjects).length,
            totalProjects: this.allProjects.size,
            projectList: Array.from(this.allProjects).sort()
        };
    }
}

module.exports = ProjectSummaryGenerator;
