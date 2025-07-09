/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.FACTDOUBLE
 * 
 *  Демонстрация использования метода FACTDOUBLE класса ApiWorksheetFunction
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
        // This example shows how to return the double factorial of a number.
        
        // How to calculate a double factorial.
        
        // Use function to calculate the double factorial of a number.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.FACTDOUBLE(123));
        
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
