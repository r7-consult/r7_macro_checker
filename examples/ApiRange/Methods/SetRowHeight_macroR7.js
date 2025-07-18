/**
 * OnlyOffice JavaScript макрос - ApiRange.SetRowHeight
 * 
 *  Демонстрация использования метода SetRowHeight класса ApiRange
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
        // This example sets the row height value.
        
        // How to set a row height of cells.
        
        // Get a range and specify its row height.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetRowHeight(32);
        
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
