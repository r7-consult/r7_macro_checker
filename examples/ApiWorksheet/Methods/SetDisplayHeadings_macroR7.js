/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.SetDisplayHeadings
 * 
 *  Демонстрация использования метода SetDisplayHeadings класса ApiWorksheet
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
        // This example specifies whether the sheet row/column headers must be displayed or not.
        
        // How to set whether sheet headings should be displayed or not.
        
        // Set a boolean value representing whether to display row/column headings or not.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("The sheet settings make it display no row/column headers");
        worksheet.SetDisplayHeadings(false);
        
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
