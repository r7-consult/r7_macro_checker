/**
 * OnlyOffice JavaScript макрос - ApiRange.SetOffset
 * 
 *  Демонстрация использования метода SetOffset класса ApiRange
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
        // This example sets the cell offset.
        
        // How to set an offset of cells.
        
        // Get a range and specify its cells offset.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B3").SetValue("Old Range");
        let range = worksheet.GetRange("B3");
        range.SetOffset(2, 2);
        range.SetValue("New Range");
        
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
