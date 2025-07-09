/**
 * OnlyOffice JavaScript макрос - ApiRange.SetFontColor
 * 
 *  Демонстрация использования метода SetFontColor класса ApiRange
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
        // This example sets the text color to the cell range.
        
        // How to color a cell text.
        
        // Get a range and apply an RGB color to its text color.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetFontColor(Api.CreateColorFromRGB(255, 111, 61));
        worksheet.GetRange("A2").SetValue("This is the text with a color set to it");
        worksheet.GetRange("A4").SetValue("This is the text with a default color");
        
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
