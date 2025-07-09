/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.AddDefName
 * 
 *  Демонстрация использования метода AddDefName класса ApiWorksheet
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
        // This example adds a new name to the worksheet.
        
        // How to change a name of the worksheet range.
        
        // Name a range from a worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        worksheet.GetRange("A3").SetValue("We defined a name 'numbers' for a range of cells A1:B1.");
        
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
