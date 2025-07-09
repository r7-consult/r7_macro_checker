/**
 * OnlyOffice JavaScript макрос - ApiRange.ForEach
 * 
 *  Демонстрация использования метода ForEach класса ApiRange
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
        // This example executes a provided function once for each cell.
        
        // How to iterate through each cell from a range.
        
        // For Each cycle implementation for ApiRange.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("C1").SetValue("3");
        let range = worksheet.GetRange("A1:C1");
        range.ForEach(function (range) {
            let value = range.GetValue();
            if (value != "1") {
                range.SetBold(true);
            }
        });
        
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
