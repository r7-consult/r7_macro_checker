/**
 * OnlyOffice JavaScript макрос - Api.RecalculateAllFormulas
 * 
 *  Демонстрация использования метода RecalculateAllFormulas класса Api
 * https://r7-consult.ru/
 */

(function() {
    'use strict';
    
    try {
        // Initialize OnlyOffice API
        const api = Api;
        if (!api) {
            throw new Error('OnlyOffice API not available');
        }
        
        // Original code enhanced with error handling:
        // This example recalculates all formulas in the active workbook.
        
        // How to recalculate all formulas in a worksheet.
        
        // Reset all values calculated by formulas.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue(1);
        worksheet.GetRange("C1").SetValue(2);
        let range = worksheet.GetRange("A1");
        range.SetValue("=SUM(B1:C1)");
        range = worksheet.GetRange("E1");
        range.SetValue("=A1+1");
        worksheet.GetRange("B1").SetValue(3);
        Api.RecalculateAllFormulas();
        worksheet.GetRange("A3").SetValue("Formulas from cells A1 and E1 were recalculated with a new value from cell C1.");
        
        // Success notification
        console.log('Macro executed successfully');
        
    } catch (error) {
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
