/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.IMSEC
 * 
 *  Демонстрация использования метода IMSEC класса ApiWorksheetFunction
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
        // This example shows how to return the secant of a complex number.
        
        // How to get the secant.
        
        // Use a function to get the secant of a complex number.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.IMSEC("-2+2.5i"));
        
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
