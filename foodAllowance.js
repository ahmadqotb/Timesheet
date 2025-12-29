const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

class FoodAllowanceCalculator {
    constructor(timesheetPath, projectPoliciesPath, employeePoliciesPath, month, year) {
        this.timesheetPath = timesheetPath;
        this.projectPoliciesPath = projectPoliciesPath;
        this.employeePoliciesPath = employeePoliciesPath;
        this.month = parseInt(month);
        this.year = parseInt(year);
        
        this.projectPolicies = {}; // {ProjectCode: {policy1: Yes/No, policy2: Yes/No}}
        this.employeePolicies = {}; // {EmployeeName: {policy: "Food Policy 1", amount: 200}}
        this.employeeData = {}; // {EmployeeName: [{date, projectCode, projectName, isAnnualLeave}]}
        this.foodAllowanceSummary = {}; // Final calculations
    }

    async process() {
        await this.loadProjectPolicies();
        await this.loadEmployeePolicies();
        await this.loadTimesheetData();
        this.calculateFoodAllowance();
    }

    async loadProjectPolicies() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(this.projectPoliciesPath);
        const worksheet = workbook.worksheets[0];

        worksheet.eachRow((row, rowIndex) => {
            if (rowIndex === 1) return; // Skip header

            const projectCode = row.getCell(1).value;
            const projectName = row.getCell(2).value;
            const location = row.getCell(3).value;
            const policy1 = row.getCell(4).value;
            const policy2 = row.getCell(5).value;

            if (projectCode) {
                this.projectPolicies[projectCode] = {
                    projectName: projectName || '',
                    location: location || '',
                    policy1: policy1 && policy1.toString().toLowerCase() === 'yes',
                    policy2: policy2 && policy2.toString().toLowerCase() === 'yes'
                };
            }
        });
    }

    async loadEmployeePolicies() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(this.employeePoliciesPath);
        const worksheet = workbook.worksheets[0];

        worksheet.eachRow((row, rowIndex) => {
            if (rowIndex === 1) return; // Skip header

            const employeeName = row.getCell(1).value;
            const amount = row.getCell(2).value;
            const policy = row.getCell(3).value;

            if (employeeName && policy) {
                this.employeePolicies[employeeName] = {
                    policy: policy.toString(),
                    amount: parseFloat(amount) || 0
                };
            }
        });
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

            // Check if Annual Leave
            const isAnnualLeave = projectName && 
                                 projectName.toLowerCase().includes('annual leave');

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

    formatDateLocal(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    calculateFoodAllowance() {
        for (const employeeName in this.employeeData) {
            // Skip if employee not in food policy list
            if (!this.employeePolicies[employeeName]) {
                continue;
            }

            const employeePolicy = this.employeePolicies[employeeName];
            const entries = this.employeeData[employeeName];

            // Group entries by date
            const dateMap = {};
            entries.forEach(entry => {
                const dateStr = this.formatDateLocal(entry.date);
                if (!dateMap[dateStr]) {
                    dateMap[dateStr] = [];
                }
                dateMap[dateStr].push(entry);
            });

            // Count eligible days
            let eligibleDays = 0;
            const dailyBreakdown = [];

            for (const dateStr in dateMap) {
                const dayEntries = dateMap[dateStr];
                
                // Check if any entry is Annual Leave
                const hasAnnualLeave = dayEntries.some(e => e.isAnnualLeave);
                if (hasAnnualLeave) {
                    // Annual Leave doesn't count
                    dailyBreakdown.push({
                        date: dateStr,
                        projects: dayEntries.map(e => e.projectCode).join(', '),
                        eligible: false,
                        reason: 'Annual Leave'
                    });
                    continue;
                }

                // Check if ANY project is eligible for employee's policy
                let isEligible = false;
                let eligibleProjects = [];

                for (const entry of dayEntries) {
                    const projectPolicy = this.projectPolicies[entry.projectCode];
                    if (projectPolicy) {
                        // Match employee policy with project eligibility
                        if (employeePolicy.policy === 'Food Policy 1' && projectPolicy.policy1) {
                            isEligible = true;
                            eligibleProjects.push(entry.projectCode);
                        } else if (employeePolicy.policy === 'Food Policy 2' && projectPolicy.policy2) {
                            isEligible = true;
                            eligibleProjects.push(entry.projectCode);
                        }
                    }
                }

                if (isEligible) {
                    eligibleDays++;
                    dailyBreakdown.push({
                        date: dateStr,
                        projects: dayEntries.map(e => e.projectCode).join(', '),
                        eligible: true,
                        eligibleProjects: eligibleProjects.join(', ')
                    });
                } else {
                    dailyBreakdown.push({
                        date: dateStr,
                        projects: dayEntries.map(e => e.projectCode).join(', '),
                        eligible: false,
                        reason: 'Project not eligible'
                    });
                }
            }

            this.foodAllowanceSummary[employeeName] = {
                policy: employeePolicy.policy,
                amountPerDay: employeePolicy.amount,
                eligibleDays: eligibleDays,
                totalAmount: eligibleDays * employeePolicy.amount,
                dailyBreakdown: dailyBreakdown
            };
        }
    }

    async generateExcelReport(outputPath) {
        const workbook = new ExcelJS.Workbook();
        
        // Format month/year for sheet name
        const monthYear = String(this.month).padStart(2, '0') + this.year;
        
        // Sheet 1: Summary
        const summarySheet = workbook.addWorksheet(`Food Allowance - ${monthYear}`);
        this.createSummarySheet(summarySheet);

        // Sheet 2: Daily Breakdown
        const detailSheet = workbook.addWorksheet('Daily Breakdown');
        this.createDetailSheet(detailSheet);

        // Sheet 3: Project Summary
        const projectSheet = workbook.addWorksheet('Project Summary');
        this.createProjectSheet(projectSheet);

        // Sheet 4: Policy Summary
        const policySheet = workbook.addWorksheet('Policy Summary');
        this.createPolicySheet(policySheet);

        await workbook.xlsx.writeFile(outputPath);
        return outputPath;
    }

    createSummarySheet(sheet) {
        // Headers
        sheet.columns = [
            { header: 'Unique Employee Name', key: 'name', width: 35 },
            { header: 'Food Allowance\nAmount / Day', key: 'amount', width: 18 },
            { header: 'Food Allowance\nDays', key: 'days', width: 15 },
            { header: 'Total', key: 'total', width: 15 }
        ];

        // Style header row
        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true, size: 11 };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        headerRow.height = 35;
        headerRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD3D3D3' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add data rows
        let grandTotal = 0;
        const sortedNames = Object.keys(this.foodAllowanceSummary).sort();

        sortedNames.forEach(employeeName => {
            const data = this.foodAllowanceSummary[employeeName];
            const row = sheet.addRow({
                name: employeeName,
                amount: data.amountPerDay,
                days: data.eligibleDays,
                total: data.totalAmount
            });

            // Format numbers with commas
            row.getCell('amount').numFmt = '#,##0';
            row.getCell('days').numFmt = '0';
            row.getCell('total').numFmt = '#,##0';

            // Borders
            row.eachCell(cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            grandTotal += data.totalAmount;
        });

        // Grand Total row
        const totalRow = sheet.addRow({
            name: 'Grand Total',
            amount: '',
            days: '',
            total: grandTotal
        });

        totalRow.font = { bold: true, size: 11 };
        totalRow.getCell('total').numFmt = '#,##0';
        totalRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFEB3B' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    }

    createDetailSheet(sheet) {
        sheet.columns = [
            { header: 'Employee Name', key: 'name', width: 35 },
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Project Code(s)', key: 'projects', width: 30 },
            { header: 'Policy', key: 'policy', width: 15 },
            { header: 'Amount/Day', key: 'amount', width: 12 },
            { header: 'Eligible?', key: 'eligible', width: 12 },
            { header: 'Reason/Projects', key: 'reason', width: 40 }
        ];

        // Style header
        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF4CAF50' }
            };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add data
        const sortedNames = Object.keys(this.foodAllowanceSummary).sort();
        sortedNames.forEach(employeeName => {
            const summary = this.foodAllowanceSummary[employeeName];
            summary.dailyBreakdown.forEach(day => {
                const row = sheet.addRow({
                    name: employeeName,
                    date: day.date,
                    projects: day.projects,
                    policy: summary.policy,
                    amount: summary.amountPerDay,
                    eligible: day.eligible ? 'Yes' : 'No',
                    reason: day.eligible ? day.eligibleProjects : day.reason
                });

                // Highlight eligible days
                if (day.eligible) {
                    row.getCell('eligible').fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFC8E6C9' }
                    };
                } else {
                    row.getCell('eligible').fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFCCCC' }
                    };
                }

                row.eachCell(cell => {
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            });
        });
    }

    createProjectSheet(sheet) {
        // Aggregate data by project
        const projectData = {};

        for (const employeeName in this.foodAllowanceSummary) {
            const summary = this.foodAllowanceSummary[employeeName];
            const policy = summary.policy;

            summary.dailyBreakdown.forEach(day => {
                if (day.eligible) {
                    const projects = day.projects.split(', ');
                    projects.forEach(projectCode => {
                        if (!projectData[projectCode]) {
                            projectData[projectCode] = {
                                policy1Days: 0,
                                policy2Days: 0,
                                policy1Amount: 0,
                                policy2Amount: 0
                            };
                        }

                        if (policy === 'Food Policy 1') {
                            projectData[projectCode].policy1Days++;
                            projectData[projectCode].policy1Amount += summary.amountPerDay;
                        } else if (policy === 'Food Policy 2') {
                            projectData[projectCode].policy2Days++;
                            projectData[projectCode].policy2Amount += summary.amountPerDay;
                        }
                    });
                }
            });
        }

        // Create sheet
        sheet.columns = [
            { header: 'Project Code', key: 'code', width: 20 },
            { header: 'Project Name', key: 'name', width: 40 },
            { header: 'Policy 1 Days', key: 'p1days', width: 15 },
            { header: 'Policy 1 Amount', key: 'p1amount', width: 18 },
            { header: 'Policy 2 Days', key: 'p2days', width: 15 },
            { header: 'Policy 2 Amount', key: 'p2amount', width: 18 },
            { header: 'Total Amount', key: 'total', width: 18 }
        ];

        // Style header
        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF2196F3' }
            };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add data
        const sortedProjects = Object.keys(projectData).sort();
        let totalP1Days = 0, totalP2Days = 0, totalP1Amount = 0, totalP2Amount = 0, grandTotal = 0;

        sortedProjects.forEach(projectCode => {
            const data = projectData[projectCode];
            const projectPolicy = this.projectPolicies[projectCode];
            const totalAmount = data.policy1Amount + data.policy2Amount;

            const row = sheet.addRow({
                code: projectCode,
                name: projectPolicy ? projectPolicy.projectName : '',
                p1days: data.policy1Days,
                p1amount: data.policy1Amount,
                p2days: data.policy2Days,
                p2amount: data.policy2Amount,
                total: totalAmount
            });

            row.getCell('p1amount').numFmt = '#,##0';
            row.getCell('p2amount').numFmt = '#,##0';
            row.getCell('total').numFmt = '#,##0';

            row.eachCell(cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            totalP1Days += data.policy1Days;
            totalP2Days += data.policy2Days;
            totalP1Amount += data.policy1Amount;
            totalP2Amount += data.policy2Amount;
            grandTotal += totalAmount;
        });

        // Total row
        const totalRow = sheet.addRow({
            code: 'Total',
            name: '',
            p1days: totalP1Days,
            p1amount: totalP1Amount,
            p2days: totalP2Days,
            p2amount: totalP2Amount,
            total: grandTotal
        });

        totalRow.font = { bold: true };
        totalRow.getCell('p1amount').numFmt = '#,##0';
        totalRow.getCell('p2amount').numFmt = '#,##0';
        totalRow.getCell('total').numFmt = '#,##0';
        totalRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFEB3B' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    }

    createPolicySheet(sheet) {
        // Aggregate by policy
        const policyData = {
            'Food Policy 1': { employees: new Set(), days: 0, amount: 0 },
            'Food Policy 2': { employees: new Set(), days: 0, amount: 0 }
        };

        for (const employeeName in this.foodAllowanceSummary) {
            const summary = this.foodAllowanceSummary[employeeName];
            const policy = summary.policy;

            if (policyData[policy]) {
                policyData[policy].employees.add(employeeName);
                policyData[policy].days += summary.eligibleDays;
                policyData[policy].amount += summary.totalAmount;
            }
        }

        // Create sheet
        sheet.columns = [
            { header: 'Policy', key: 'policy', width: 20 },
            { header: 'Total Employees', key: 'employees', width: 18 },
            { header: 'Total Days', key: 'days', width: 15 },
            { header: 'Total Amount', key: 'amount', width: 18 }
        ];

        // Style header
        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF9800' }
            };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Add data
        let totalEmployees = 0, totalDays = 0, totalAmount = 0;

        ['Food Policy 1', 'Food Policy 2'].forEach(policy => {
            const data = policyData[policy];
            const row = sheet.addRow({
                policy: policy,
                employees: data.employees.size,
                days: data.days,
                amount: data.amount
            });

            row.getCell('amount').numFmt = '#,##0';
            row.eachCell(cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            totalEmployees += data.employees.size;
            totalDays += data.days;
            totalAmount += data.amount;
        });

        // Total row
        const totalRow = sheet.addRow({
            policy: 'Total',
            employees: totalEmployees,
            days: totalDays,
            amount: totalAmount
        });

        totalRow.font = { bold: true };
        totalRow.getCell('amount').numFmt = '#,##0';
        totalRow.eachCell(cell => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFEB3B' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    }
}

module.exports = FoodAllowanceCalculator;
