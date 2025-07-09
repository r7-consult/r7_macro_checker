/**
 * OnlyOffice JavaScript макрос - ApiRange.GetValue
 * 
 *  Демонстрация использования метода GetValue класса ApiRange
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
        // This example shows how to get a value of the specified range.
        
        // How to get a cell value.
        
        // Get a range, get its value and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let value = worksheet.GetRange("A1").GetValue();
        worksheet.GetRange("A3").SetValue("Value of the cell A1: ");
        worksheet.GetRange("B3").SetValue(value);
        
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
