/**
 * OnlyOffice JavaScript макрос - ApiRange.Insert
 * 
 *  Демонстрация использования метода Insert класса ApiRange
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
        // This example inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.
        
        // How to insert a range or a cell into a worksheet.
        
        // Insert a range or a cell into a worksheet specifying its shift direction.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B4").SetValue("1");
        worksheet.GetRange("C4").SetValue("2");
        worksheet.GetRange("D4").SetValue("3");
        worksheet.GetRange("C5").SetValue("5");
        let range = worksheet.GetRange("C4");
        range.Insert("down");
        
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
