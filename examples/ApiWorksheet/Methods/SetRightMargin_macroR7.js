/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetRightMargin
 * 
 *  Демонстрация использования метода SetRightMargin класса ApiWorksheet
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
        // This example sets the right margin of the sheet.
        
        // How to set margin of the right side.
        
        // Resize the right margin of the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetRightMargin(20.8);
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
