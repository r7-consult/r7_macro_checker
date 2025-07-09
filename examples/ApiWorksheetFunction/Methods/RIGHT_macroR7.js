/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.RIGHT
 * 
 *  Демонстрация использования метода RIGHT класса ApiWorksheetFunction
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
        // This example shows how to the specified number of characters from the end of a text string.
        
        // How to get characters from the end of a string.
        
        // Use a function to return n characters from the right part of a text.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.RIGHT("Online Office", 2));
        
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
