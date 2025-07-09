/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetName
 * 
 *  Демонстрация использования метода GetName класса ApiWorksheet
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
        // This example shows how to get a sheet name.
        
        // How to get name of the sheet.
        
        // Get a sheet name.
        
        let worksheet = Api.GetActiveSheet();
        let name = worksheet.GetName();
        worksheet.GetRange("A1").SetValue("Name: ");
        worksheet.GetRange("B1").SetValue(name);
        
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
