const XLSX = require('xlsx');
const fs = require('fs');

// Function to analyze the Excel data
function analyzeExcelData(file_path) {
    // Load the spreadsheet
    const workbook = XLSX.readFile(file_path);
    const sheet_name = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheet_name];

    // Convert the sheet data to JSON
    const data = XLSX.utils.sheet_to_json(sheet);

    // Step 1: Filter employees who have worked for 7 consecutive days
    const consecutive_days_threshold = 7;
    const employees_7_consecutive_days = data.filter((employee, index, array) => {
        if (index >= consecutive_days_threshold) {
            const current = new Date(employee['Pay Cycle Start Date'] + ' ' + employee['Time']);
            const previous = new Date(array[index - consecutive_days_threshold]['Pay Cycle Start Date'] + ' ' + array[index - consecutive_days_threshold]['Time']);
            const diffInDays = (current - previous) / (1000 * 60 * 60 * 24);
            return diffInDays === consecutive_days_threshold;
        }
        return false;
    });

    // Step 2: Filter employees with less than 10 hours between shifts but greater than 1 hour
    const min_hours_between_shifts = 1;
    const max_hours_between_shifts = 10;
    const employees_less_than_10_hours_between_shifts = data.filter((employee, index, array) => {
        if (index > 0) {
            const current = new Date(employee['Pay Cycle Start Date'] + ' ' + employee['Time']);
            const previous = new Date(array[index - 1]['Pay Cycle Start Date'] + ' ' + array[index - 1]['Time']);
            const diffInHours = (current - previous) / (1000 * 60 * 60);
            return diffInHours > min_hours_between_shifts && diffInHours < max_hours_between_shifts;
        }
        return false;
    });

    // Step 3: Filter employees who have worked for more than 14 hours in a single shift
    const max_hours_single_shift = 14;
    const employees_more_than_14_hours = data.filter(employee => parseFloat(employee['Timecard Hours (as Time)']) > max_hours_single_shift);

    // Print the results to the console
    console.log("Employees with 7 consecutive days:");
    console.table(employees_7_consecutive_days.map(employee => ({ 'Position ID': employee['Position ID'], 'Employee Name': employee['Employee Name'] })));

    console.log("\nEmployees with less than 10 hours between shifts but greater than 1 hour:");
    console.table(employees_less_than_10_hours_between_shifts.map(employee => ({ 'Position ID': employee['Position ID'], 'Employee Name': employee['Employee Name'] })));

    console.log("\nEmployees who have worked for more than 14 hours in a single shift:");
    console.table(employees_more_than_14_hours.map(employee => ({ 'Position ID': employee['Position ID'], 'Employee Name': employee['Employee Name'] })));

    // Save the results to a text file
    const output_path = 'output.txt';
    fs.writeFileSync(output_path, '');

    fs.appendFileSync(output_path, "Employees with 7 consecutive days:\n");
    fs.appendFileSync(output_path, JSON.stringify(employees_7_consecutive_days.map(employee => ({ 'Position ID': employee['Position ID'], 'Employee Name': employee['Employee Name'] })), null, 2) + "\n\n");

    fs.appendFileSync(output_path, "Employees with less than 10 hours between shifts but greater than 1 hour:\n");
    fs.appendFileSync(output_path, JSON.stringify(employees_less_than_10_hours_between_shifts.map(employee => ({ 'Position ID': employee['Position ID'], 'Employee Name': employee['Employee Name'] })), null, 2) + "\n\n");

    fs.appendFileSync(output_path, "Employees who have worked for more than 14 hours in a single shift:\n");
    fs.appendFileSync(output_path, JSON.stringify(employees_more_than_14_hours.map(employee => ({ 'Position ID': employee['Position ID'], 'Employee Name': employee['Employee Name'] })), null, 2) + "\n");
}

// Execute the analysis with the provided file path
const file_path = 'file.xlsx'; // Relative path from the script file
analyzeExcelData(file_path);
