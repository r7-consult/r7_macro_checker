/**
 * OnlyOffice JavaScript макрос - ApiRange.End
 * 
 *  Демонстрация использования метода End класса ApiRange
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
        // This example shows how to get a Range object that represents the end in the specified direction in the specified range.
        
        // Get a left end part of a range and fill it with color.
        
        // Get a specified direction end of a range.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("C4:D5");
        range.End("xlToLeft").SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
