/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.CHOOSE
 * 
 *  Демонстрация использования метода CHOOSE класса ApiWorksheetFunction
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
        // This example shows how to choose a value or action to perform from a list of values, based on an index number.
        
        // How to choose a value or action to perform from a list of values, based on an index number.
        
        // Use function to choose a value or action to perform from a list of values, based on an index number.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.CHOOSE(2, 3, 4, 89, 76, 0));
        
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
