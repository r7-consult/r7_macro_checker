/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetLeftMargin
 * 
 *  Демонстрация использования метода SetLeftMargin класса ApiWorksheet
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
        // This example sets the left margin of the sheet.
        
        // How to set margin of the left side.
        
        // Resize the left margin of the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetLeftMargin(20.8);
        let leftMargin = worksheet.GetLeftMargin();
        worksheet.GetRange("A1").SetValue("Left margin: " + leftMargin + " mm");
        
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
