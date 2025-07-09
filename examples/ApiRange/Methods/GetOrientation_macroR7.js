/**
 * OnlyOffice JavaScript макрос - ApiRange.GetOrientation
 * 
 *  Демонстрация использования метода GetOrientation класса ApiRange
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
        // This example shows how to get the range angle.
        
        // How to find out cell orientation of a range.
        
        // Get a range, get its orientation (upward, downward, etc.) and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        let range = worksheet.GetRange("A1:B1");
        range.SetOrientation("xlUpward");
        let orientation = range.GetOrientation();
        worksheet.GetRange("A3").SetValue("Orientation: " + orientation);
        
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
