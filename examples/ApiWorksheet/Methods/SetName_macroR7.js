/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetName
 * 
 *  Демонстрация использования метода SetName класса ApiWorksheet
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
        // This example sets a name to the active sheet.
        
        // How to set name of the sheet.
        
        // Rename the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetName("sheet 1");
        let name = worksheet.GetName();
        worksheet.GetRange("A1").SetValue("Worksheet name: ");
        worksheet.GetRange("A1").AutoFit(false, true);
        worksheet.GetRange("B1").SetValue(name);
        
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
