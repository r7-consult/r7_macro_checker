/**
 * OnlyOffice JavaScript макрос - ApiRange.SetAlignVertical
 * 
 *  Демонстрация использования метода SetAlignVertical класса ApiRange
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
        // This example sets the vertical alignment of the text in the cell range.
        
        // How to change the vertical alignment of the cell content.
        
        // Change the vertical alignment of the ApiRange content to distributed.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1:D5");
        worksheet.GetRange("A2").SetValue("This is just a sample text distributed in the A2 cell.");
        range.SetAlignVertical("distributed");
        
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
