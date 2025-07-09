/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetPageOrientation
 * 
 *  Демонстрация использования метода GetPageOrientation класса ApiWorksheet
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
        // This example shows how to get the page orientation.
        
        // How to get orientation of the sheet.
        
        // Get a sheet orientation.
        
        let worksheet = Api.GetActiveSheet();
        let pageOrientation = worksheet.GetPageOrientation();
        worksheet.GetRange("A1").SetValue("Page orientation: ");
        worksheet.GetRange("C1").SetValue(pageOrientation);
        
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
