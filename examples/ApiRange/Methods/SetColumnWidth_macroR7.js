/**
 * OnlyOffice JavaScript макрос - ApiRange.SetColumnWidth
 * 
 *  Демонстрация использования метода SetColumnWidth класса ApiRange
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
        // This example sets the width of all the columns in the range.
        
        // How to make a cell column wider.
        
        // Get a range and set its column width.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetColumnWidth(20);
        
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
