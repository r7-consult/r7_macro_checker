/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetCols
 * 
 *  Демонстрация использования метода GetCols класса ApiWorksheet
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
        // This example shows how to get the ApiRange object that represents all the cells on the columns range.
        
        // How to get all column cells.
        
        // Get all column cells from the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let cols = worksheet.GetCols("A1:C1");
        cols.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
