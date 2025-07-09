/**
 * OnlyOffice JavaScript макрос - ApiRange.GetRowHeight
 * 
 *  Демонстрация использования метода GetRowHeight класса ApiRange
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
        // This example shows how to get the row height value.
        
        // How to get a cell row height.
        
        // Get a range and display its row height in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let height = worksheet.GetRange("A1").GetRowHeight();
        worksheet.GetRange("A1").SetValue("Height: ");
        worksheet.GetRange("B1").SetValue(height);
        
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
