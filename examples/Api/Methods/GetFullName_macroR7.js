/**
 * OnlyOffice JavaScript макрос - Api.GetFullName
 * 
 *  Демонстрация использования метода GetFullName класса Api
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
        // This example shows how to get the full name of the currently opened file.
        
        // How to get a full name of the file.
        
        // Insert a full name of the file into a cell.
        
        let worksheet = Api.GetActiveSheet();
        let name = Api.GetFullName();
        worksheet.GetRange("B1").SetValue("File name: " + name);
        
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
