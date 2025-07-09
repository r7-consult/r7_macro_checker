/**
 * OnlyOffice JavaScript макрос - Api.SetReferenceStyle
 * 
 *  Демонстрация использования метода SetReferenceStyle класса Api
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
        // This example sets reference style.
        
        // How to set a style of a reference.
        
        // Set reference style using ID.
        
        let worksheet = Api.GetActiveSheet();
        Api.SetReferenceStyle("xlR1C1");
        worksheet.GetRange("A1").SetValue(Api.GetReferenceStyle());
        
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
