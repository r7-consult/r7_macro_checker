/**
 * OnlyOffice JavaScript макрос - ApiRange.GetText
 * 
 *  Демонстрация использования метода GetText класса ApiRange
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
        // This example shows how to get the text of the specified range.
        
        // How to get a cell raw text value.
        
        // Get a range, get its text value and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("text1");
        worksheet.GetRange("B1").SetValue("text2");
        worksheet.GetRange("C1").SetValue("text3");
        let range = worksheet.GetRange("A1:C1");
        let text = range.GetText();
        worksheet.GetRange("A3").SetValue("Text from the cell A1: " + text);
        
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
