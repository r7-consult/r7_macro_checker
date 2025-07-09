/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.DEC2BIN
 * 
 *  Демонстрация использования метода DEC2BIN класса ApiWorksheetFunction
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
        // This example shows how to convert a decimal number to binary. 
        
        // How to get decimal number from binary.
        
        // Use function to convert a decimal number to binary.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.DEC2BIN(-100));
        
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
