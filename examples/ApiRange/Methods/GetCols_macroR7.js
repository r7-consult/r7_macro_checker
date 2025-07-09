/**
 * OnlyOffice JavaScript макрос - ApiRange.GetCols
 * 
 *  Демонстрация использования метода GetCols класса ApiRange
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
        // This example shows how to get a Range object that represents the columns in the specified range.
        
        // How to get columns from a range.
        
        // Get a range, get its first two columns and fill them with a color.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1:C3");
        range.GetCols(2).SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
