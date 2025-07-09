/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.SQRTPI
 * 
 *  Демонстрация использования метода SQRTPI класса ApiWorksheetFunction
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
        // This example shows how to return the square root of (number * pi).
        
        // How to calculate the square root of a number multiplied by aa constant Pi.
        
        // Use a function to get the square root of (number * pi).
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.SQRTPI(5));
        
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
