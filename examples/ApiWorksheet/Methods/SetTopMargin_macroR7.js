/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetTopMargin
 * 
 *  Демонстрация использования метода SetTopMargin класса ApiWorksheet
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
        // This example sets the top margin of the sheet.
        
        // How to set margin of the top side.
        
        // Resize the top margin of the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetTopMargin(25.1);
        let topMargin = worksheet.GetTopMargin();
        worksheet.GetRange("A1").SetValue("Top margin: " + topMargin + " mm");
        
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
