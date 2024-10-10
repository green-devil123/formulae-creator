
// Validation of formulas (example implementation)
const validateFormulas = (worksheet) => {
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            if (cell.formula) {
                // Check if the formula references invalid cells or operations
                const formula = cell.formula.toString();
                if (!/^[A-Z]+\d+/.test(formula) && !formula.includes('VALUE') && !formula.includes('MID')) {
                    console.warn(`Warning: Cell ${cell.address} has an invalid formula: ${formula}`);
                }else{
                    console.warn(`${cell.address} has an valid formula: ${formula}`);
                }
            }
        });
    });
};

module.exports = { validateFormulas };
