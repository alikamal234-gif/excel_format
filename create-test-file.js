const XLSX = require('xlsx');
const path = require('path');

// Create a sample workbook for testing
const wb = XLSX.utils.book_new();

// Sheet 1: Employee Data
const sheet1Data = [
  ['ID', 'Name', 'Department', 'Salary', 'Start Date', 'Email'],
  [1, 'Ahmed Bennani', 'Engineering', 85000, '2022-01-15', 'ahmed@company.com'],
  [2, 'Fatima Zahra', 'Marketing', 72000, '2021-06-20', 'fatima@company.com'],
  [3, 'Youssef El Amrani', 'Engineering', 90000, '2020-03-10', 'youssef@company.com'],
  [4, 'Khadija Mouline', 'HR', 68000, '2023-02-01', 'khadija@company.com'],
  [5, 'Omar Tazi', 'Sales', 78000, '2021-11-05', 'omar@company.com'],
  [6, 'Sara Idrissi', 'Engineering', 95000, '2019-08-22', 'sara@company.com'],
  [7, 'Hamza Berrada', 'Marketing', 71000, '2022-04-18', 'hamza@company.com'],
  [8, 'Leila Senhaji', 'Finance', 82000, '2020-12-01', 'leila@company.com'],
  [9, 'Amine Chraibi', 'Sales', 76000, '2023-07-15', 'amine@company.com'],
  [10, 'Nadia Fassi', 'HR', 70000, '2021-09-30', 'nadia@company.com'],
];
const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);
ws1['!cols'] = [
  { wch: 5 }, { wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 12 }, { wch: 25 },
];
XLSX.utils.book_append_sheet(wb, ws1, 'Employees');

// Sheet 2: Sales Data
const sheet2Data = [
  ['Month', 'Product A', 'Product B', 'Product C', 'Total'],
  ['January', 15000, 22000, 18000, 55000],
  ['February', 18000, 19000, 21000, 58000],
  ['March', 22000, 25000, 17000, 64000],
  ['April', 20000, 28000, 23000, 71000],
  ['May', 25000, 24000, 26000, 75000],
  ['June', 19000, 30000, 22000, 71000],
];
const ws2 = XLSX.utils.aoa_to_sheet(sheet2Data);
XLSX.utils.book_append_sheet(wb, ws2, 'Sales');

// Sheet 3: Project Status
const sheet3Data = [
  ['Project', 'Status', 'Progress %', 'Lead', 'Deadline'],
  ['Website Redesign', 'In Progress', 75, 'Ahmed', '2024-03-15'],
  ['Mobile App v2', 'Planning', 15, 'Sara', '2024-06-01'],
  ['API Migration', 'Completed', 100, 'Youssef', '2024-01-30'],
  ['Data Pipeline', 'In Progress', 45, 'Leila', '2024-04-20'],
  ['UI Components', 'In Progress', 60, 'Hamza', '2024-03-01'],
];
const ws3 = XLSX.utils.aoa_to_sheet(sheet3Data);
XLSX.utils.book_append_sheet(wb, ws3, 'Projects');

// Write to file
const outputPath = path.join(__dirname, 'test-data.xlsx');
XLSX.writeFile(wb, outputPath);
console.log('Test file created at:', outputPath);
