/**
 * OnlyOffice JavaScript макрос - ApiRange.SetFormulaArray
 * 
 *  Демонстрация использования метода SetFormulaArray класса ApiRange
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
        // This example sets the array formula of a range.
        
        // How to set the array formula value.
        
        // Set the array formula.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1:C3").SetFormulaArray("={1,2,3}");
        
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
