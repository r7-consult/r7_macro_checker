/**
 * OnlyOffice JavaScript макрос - Api.GetDefName
 * 
 *  Демонстрация использования метода GetDefName класса Api
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
        // This example shows how to get the ApiName object by the range name.
        
        // How to work with named ranges in a spreadsheet using the API.
        
        // Get name of an object using a range name. 
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        let defName = Api.GetDefName("numbers");
        worksheet.GetRange("A3").SetValue("DefName: " + defName.GetName());
        
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
