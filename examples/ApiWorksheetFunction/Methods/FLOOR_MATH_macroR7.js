/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.FLOOR_MATH
 * 
 *  Демонстрация использования метода FLOOR_MATH класса ApiWorksheetFunction
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
        // This example shows how to round a number down, to the nearest integer or to the nearest multiple of significance.
        
        // How to round a number down to the nearest integer.
        
        // Use function to round down a number with specified decimal points.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.FLOOR_MATH(-5.5, 2, 1));
        
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
