/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetRowHeight
 * 
 *  Демонстрация использования метода SetRowHeight класса ApiWorksheet
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
        // This example sets the height of the specified row measured in points.
        
        // How to resize the height of the row.
        
        // Set a row height.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetRowHeight(0, 30);
        
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
