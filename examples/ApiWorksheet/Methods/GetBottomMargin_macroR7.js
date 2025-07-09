/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetBottomMargin
 * 
 *  Демонстрация использования метода GetBottomMargin класса ApiWorksheet
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
        // This example shows how to get the bottom margin of the sheet.
        
        // How to get margin of the bottom.
        
        // Get the size of the bottom margin of the sheet.
        
        let worksheet = Api.GetActiveSheet();
        let bottomMargin = worksheet.GetBottomMargin();
        worksheet.GetRange("A1").SetValue("Bottom margin: " + bottomMargin + " mm");
        
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
