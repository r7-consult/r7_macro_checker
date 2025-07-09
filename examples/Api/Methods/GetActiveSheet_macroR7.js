/**
 * OnlyOffice JavaScript макрос - Api.GetActiveSheet
 * 
 *  Демонстрация использования метода GetActiveSheet класса Api
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
        // This example shows how to get an object that represents the active sheet.
        
        // How to get an active sheet object.
        
        // Get worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("B2").SetValue("2");
        worksheet.GetRange("A3").SetValue("2x2=");
        worksheet.GetRange("B3").SetValue("=B1*B2");
        
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
