/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.Paste
 * 
 *  Демонстрация использования метода Paste класса ApiWorksheet
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
        // This example shows how to paste date from clipboard.
        
        // How to paste a copied or cut data from the clipboard.
        
        // Paste to the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.Paste();
        
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
