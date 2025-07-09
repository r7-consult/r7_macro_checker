/**
 * OnlyOffice JavaScript макрос - ApiRange.SetBorders
 * 
 *  Демонстрация использования метода SetBorders класса ApiRange
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
        // This example sets the border to the cell with the parameters specified.
        
        // How to set the thick bottom border to a cell.
        
        // Get a range and set its border specifying its side, type and color.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetColumnWidth(0, 50);
        worksheet.GetRange("A2").SetBorders("Bottom", "Thick", Api.CreateColorFromRGB(255, 111, 61));
        worksheet.GetRange("A2").SetValue("This is a cell with a bottom border");
        
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
