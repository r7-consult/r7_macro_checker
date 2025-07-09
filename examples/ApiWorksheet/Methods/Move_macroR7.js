/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.Move
 * 
 *  Демонстрация использования метода Move класса ApiWorksheet
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
        // This example moves the sheet to another location in the workbook.
        
        // How to change an order of the sheet.
        
        // Move a sheet.
        
        let sheet1 = Api.GetActiveSheet();
        Api.AddSheet("Sheet2");
        let sheet2 = Api.GetActiveSheet();
        sheet2.Move(sheet1);
        
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
