/**
 * OnlyOffice JavaScript макрос - Api.Intersect
 * 
 *  Демонстрация использования метода Intersect класса Api
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
        // This example shows how to get the ApiRange object that represents the rectangular intersection of two or more ranges.
        
        // How to find intersection of two ranges and highlight it.
        
        // Find common cells of two ranges and fill them with a color.
        
        let worksheet = Api.GetActiveSheet();
        let range1 = worksheet.GetRange("A1:C5");
        let range2 = worksheet.GetRange("B2:B4");
        let range = Api.Intersect(range1, range2);
        range.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
