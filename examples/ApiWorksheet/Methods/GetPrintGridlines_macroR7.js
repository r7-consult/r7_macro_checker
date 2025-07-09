/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetPrintGridlines
 * 
 *  Демонстрация использования метода GetPrintGridlines класса ApiWorksheet
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
        // This example shows how to get the page PrintGridlines property which specifies whether the sheet gridlines must be printed or not.
        
        // How to find out whether sheet gridlines should be printed or not.
        
        // Get a boolean value representing whether to print gridlines or not.
        
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
