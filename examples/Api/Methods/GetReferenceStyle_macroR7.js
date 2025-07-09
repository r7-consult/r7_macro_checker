/**
 * OnlyOffice JavaScript макрос - Api.GetReferenceStyle
 * 
 *  Демонстрация использования метода GetReferenceStyle класса Api
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
        // This example gets reference style.
        
        // Get style of a reference.
        
        // Insert a reference style into the cell.
        
        let worksheet = Api.GetActiveSheet();
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
