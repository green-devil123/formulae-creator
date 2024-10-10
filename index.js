const ExcelJS = require('exceljs');
const { validateFormulas } = require('./validateFormulas');

const workbook = new ExcelJS.Workbook();

// Read the source Excel file
workbook.xlsx.readFile('source.xlsx').then(() => {
    
    // Select the first worksheet
    const worksheet = workbook.getWorksheet(1);

    // Insert formulas and apply font formatting
    const insertFormulaWithFont = (cellAddress, formula) => {
        const cell = worksheet.getCell(cellAddress);
        cell.value = { formula: formula };
    };

    // Determine the last column (for example, let's say you want to apply formulas to columns B to E)
    const startColumn = 2; // Column B
    const endColumn = 5;   // Column E (you can extend this as needed)

    // Insert formulas in each column
    for (let col = startColumn; col <= endColumn; col++) {
        const columnLetter = String.fromCharCode(64 + col); // Convert column index to letter

        // Insert formulas in each column

        // Gross Profit
        insertFormulaWithFont(`${columnLetter}5`, `${columnLetter}3 - ${columnLetter}4`);  
        
        // Total Operating Expenses
        insertFormulaWithFont(`${columnLetter}10`, `${columnLetter}7 + ${columnLetter}8 + ${columnLetter}9`);  
        
        // Operating Income
        insertFormulaWithFont(`${columnLetter}11`, `${columnLetter}5 - ${columnLetter}10`);  
        
        // formula logic for Other income (expense), net
        const formulaForB15 = `${columnLetter}12 - VALUE(MID(${columnLetter}13, 2, LEN(${columnLetter}13)-2)) + ${columnLetter}14`;
        insertFormulaWithFont(`${columnLetter}15`, formulaForB15);  // Insert formula for Other income
        
        // Income before income tax
        insertFormulaWithFont(`${columnLetter}16`, `${columnLetter}11 + ${columnLetter}15`);  
        
        // Net income
        insertFormulaWithFont(`${columnLetter}18`, `${columnLetter}16 - ${columnLetter}17`);  
    }

    // Validate the formulas after inserting them
    validateFormulas(worksheet);

    // Save the updated workbook with fonts, formulas, and formatted dates
    return workbook.xlsx.writeFile('target_with_formulas.xlsx');

}).then(() => {
    console.log('Formulas and fonts have been applied, and the file has been saved.');
}).catch(err => {
    console.error('Error:', err);
});
