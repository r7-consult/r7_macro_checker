/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetPrintHeadings
 * 
 *  Демонстрация использования метода SetPrintHeadings класса ApiWorksheet
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
        // This example specifies whether the current sheet row/column headers must be printed or not.
        
        // How to set whether sheet headings should be printed or not.
        
        // Set a boolean value representing whether to print row/column headings or not.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetPrintHeadings(true);
        worksheet.GetRange("A1").SetValue("Row and column headings will be printed with this page: " + worksheet.GetPrintHeadings());
        
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
