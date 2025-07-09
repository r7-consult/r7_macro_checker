/**
 * OnlyOffice JavaScript макрос - Api.GetSheet
 * 
 *  Демонстрация использования метода GetSheet класса Api
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
        // This example shows how to get an object that represents a sheet.
        
        // How to get a sheet knowing its name.
        
        // Find and get a sheet object by its name.
        
        let worksheet = Api.GetSheet("Sheet1");
        worksheet.GetRange("A1").SetValue("This is a sample text on 'Sheet1'.");
        
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
