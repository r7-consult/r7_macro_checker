/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetRightMargin
 * 
 *  Демонстрация использования метода GetRightMargin класса ApiWorksheet
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
        // This example shows how to get the right margin of the sheet.
        
        // How to get margin of the sheet's right side.
        
        // Get the size of the right margin of the sheet.
        
        let worksheet = Api.GetActiveSheet();
        let rightMargin = worksheet.GetRightMargin();
        worksheet.GetRange("A1").SetValue("Right margin: " + rightMargin + " mm");
        
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
