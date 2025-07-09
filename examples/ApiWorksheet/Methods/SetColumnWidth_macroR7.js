/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetColumnWidth
 * 
 *  Демонстрация использования метода SetColumnWidth класса ApiWorksheet
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
        // This example sets the width of the specified column.
        
        // How to set a column width.
        
        // Resize column width.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetColumnWidth(0, 10);
        worksheet.SetColumnWidth(1, 20);
        
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
