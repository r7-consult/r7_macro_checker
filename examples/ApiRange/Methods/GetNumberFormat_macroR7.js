/**
 * OnlyOffice JavaScript макрос - ApiRange.GetNumberFormat
 * 
 *  Демонстрация использования метода GetNumberFormat класса ApiRange
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
        // This example shows how to get a value that represents the format code for the current range.
        
        // How to find out a number format of a range.
        
        // Get a range, get its cell number format and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B2");
        range.SetValue(3);
        let format = range.GetNumberFormat();
        worksheet.GetRange("B3").SetValue("Number format: " + format);
        
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
