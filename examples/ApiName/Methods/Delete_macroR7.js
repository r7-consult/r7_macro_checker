/**
 * OnlyOffice JavaScript макрос - ApiName.Delete
 * 
 *  Демонстрация использования метода Delete класса ApiName
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
        // This example deletes the DefName object.
        
        // How to remove custom DefName from a worksheet.
        
        // Delete previously added DefName. 
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        let defName = Api.GetDefName("numbers");
        defName.Delete();
        worksheet.GetRange("A3").SetValue("The name 'numbers' of the range A1:B1 was deleted.");
        
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
