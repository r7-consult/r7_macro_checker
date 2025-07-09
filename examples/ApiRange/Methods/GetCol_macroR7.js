/**
 * OnlyOffice JavaScript макрос - ApiRange.GetCol
 * 
 *  Демонстрация использования метода GetCol класса ApiRange
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
        // This example shows how to get a column number for the selected cell.
        
        // How to get a cell column index.
        
        // Get a range and display its column number.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("D9").GetCol();
        worksheet.GetRange("A2").SetValue(range.toString());
        
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
