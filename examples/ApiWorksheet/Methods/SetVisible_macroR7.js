/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetVisible
 * 
 *  Демонстрация использования метода SetVisible класса ApiWorksheet
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
        // This example sets the state of sheet visibility.
        
        // How to set visibility of the sheet.
        
        // Make a sheet visible or not.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetVisible(true);
        worksheet.GetRange("A1").SetValue("The current worksheet is visible.");
        
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
