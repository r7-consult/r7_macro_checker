/**
 * OnlyOffice JavaScript макрос - Api.GetSheets
 * 
 *  Демонстрация использования метода GetSheets класса Api
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
        // This example shows how to get a sheet collection that represents all the sheets in the active workbook.
        
        // Get all sheets as an array.
        
        // How to get array of sheets.
        
        Api.AddSheet("new_sheet_name");
        let sheets = Api.GetSheets();
        let sheetName1 = sheets[0].GetName();
        let sheetName2 = sheets[1].GetName();
        sheets[1].GetRange("A1").SetValue(sheetName1);
        sheets[1].GetRange("A2").SetValue(sheetName2);
        
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
