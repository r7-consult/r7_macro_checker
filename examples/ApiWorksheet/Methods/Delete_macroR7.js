/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.Delete
 * 
 *  Демонстрация использования метода Delete класса ApiWorksheet
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
        // This example deletes the worksheet.
        
        // How to delete sheets.
        
        // Remove a worksheet.
        
        Api.AddSheet("New sheet");
        let sheet = Api.GetActiveSheet();
        sheet.Delete();
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A3").SetValue("This method just deleted the second sheet from this spreadsheet.");
        
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
