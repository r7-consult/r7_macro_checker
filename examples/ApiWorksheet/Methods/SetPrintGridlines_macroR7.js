/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetPrintGridlines
 * 
 *  Демонстрация использования метода SetPrintGridlines класса ApiWorksheet
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
        // This example specifies whether the sheet gridlines must be printed or not.
        
        // How to set whether sheet gridlines should be printed or not.
        
        // Set a boolean value representing whether to print gridlines or not.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetPrintGridlines(true);
        worksheet.GetRange("A1").SetValue("Gridlines of cells will be printed on this page: " + worksheet.GetPrintGridlines());
        
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
