/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetActive
 * 
 *  Демонстрация использования метода SetActive класса ApiWorksheet
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
        // This example makes the sheet active.
        
        // How to set an active sheet.
        
        // Set a current sheet active.
        
        Api.AddSheet("New_sheet");
        let sheet = Api.GetSheet("New_sheet");
        sheet.SetActive();
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("The current sheet is active.");
        
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
