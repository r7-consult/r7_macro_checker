/**
 * OnlyOffice JavaScript макрос - ApiRange.GetValue2
 * 
 *  Демонстрация использования метода GetValue2 класса ApiRange
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
        // This example shows how to get the value without format of the specified range.
        
        // How to get a cell raw value.
        
        // Get a range, get its raw value without format and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let format = Api.Format("123456", "$#,##0");
        let range = worksheet.GetRange("A1");
        range.SetValue(format);
        let value2 = range.GetValue2();
        worksheet.GetRange("A3").SetValue("Value of the cell A1 without format: " + value2);
        
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
