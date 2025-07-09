/**
 * OnlyOffice JavaScript макрос - ApiRange.UnMerge
 * 
 *  Демонстрация использования метода UnMerge класса ApiRange
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
        // This example splits the selected merged cell range into the single cells.
        
        // How to unmerge a range of cells.
        
        // Get a range and split its merged cells.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A3:E8").Merge(true);
        worksheet.GetRange("A5:E5").UnMerge();
        
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
