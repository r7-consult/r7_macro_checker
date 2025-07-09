/**
 * OnlyOffice JavaScript макрос - Api.SetLocale
 * 
 *  Демонстрация использования метода SetLocale класса Api
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
        // This example sets a locale to the document.
        
        // How to set a region to the document.
        
        // Set or change the locale of the document.
        
        let worksheet = Api.GetActiveSheet();
        Api.SetLocale("en-CA");
        worksheet.GetRange("A1").SetValue("A sample spreadsheet with the language set to English (Canada).");
        
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
