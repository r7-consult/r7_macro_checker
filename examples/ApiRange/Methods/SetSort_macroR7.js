/**
 * OnlyOffice JavaScript макрос - ApiRange.SetSort
 * 
 *  Демонстрация использования метода SetSort класса ApiRange
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
        // This example sorts the cells in the given range by the parameters specified in the request.
        
        // How to sort values of cells specifying the order.
        
        // Get a range and sort its values.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue(2016);
        worksheet.GetRange("A2").SetValue(2015);
        worksheet.GetRange("A3").SetValue(2018);
        worksheet.GetRange("A4").SetValue(2014);
        worksheet.GetRange("A5").SetValue(2010);
        worksheet.GetRange("B1").SetValue(150);
        worksheet.GetRange("B2").SetValue(200);
        worksheet.GetRange("B3").SetValue(260);
        worksheet.GetRange("B4").SetValue(120);
        worksheet.GetRange("B5").SetValue(100);
        worksheet.GetRange("C1").SetValue("C");
        worksheet.GetRange("C2").SetValue("B");
        worksheet.GetRange("C3").SetValue("A");
        worksheet.GetRange("C4").SetValue("G");
        worksheet.GetRange("C5").SetValue("E");
        worksheet.GetRange("A1:C5").SetSort("A1:A5", "xlAscending", "B1:B5", "xlDescending", "C1:C5", "xlAscending", "xlYes", "xlSortColumns");
        
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
