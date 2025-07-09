/**
 * OnlyOffice JavaScript макрос - ApiRange.SetHidden
 * 
 *  Демонстрация использования метода SetHidden класса ApiRange
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
        // This example sets the value hiding property.
        
        // How to hide cells from a range.
        
        // Get a range and make specified cells invisible.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRows("1:3");
        range.SetHidden(true);
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("C1").SetValue("3");
        let hidden = range.GetHidden();
        worksheet.GetRange("A4").SetValue("The values from A1:C1 are hidden: " + hidden);
        
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
