/**
 * OnlyOffice JavaScript макрос - Api.Save
 * 
 *  Демонстрация использования метода Save класса Api
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
        // This example saves changes to the specified document.
        
        // How to save changes of the spreadsheets.
        
        // Save all applied changes.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("This sample text is saved to the worksheet.");
        Api.Save();
        
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
