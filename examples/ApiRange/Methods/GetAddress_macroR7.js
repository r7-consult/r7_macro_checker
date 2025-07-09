/**
 * OnlyOffice JavaScript макрос - ApiRange.GetAddress
 * 
 *  Демонстрация использования метода GetAddress класса ApiRange
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
        // This example shows how to get the range address.
        
        // How to get an address of a range.
        
        // Get an address of one range and set it for another one.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        let address = worksheet.GetRange("A1").GetAddress(true, true, "xlA1", false);
        worksheet.GetRange("A3").SetValue("Address: ");
        worksheet.GetRange("B3").SetValue(address);
        
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
