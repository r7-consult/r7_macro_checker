/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.WORKDAY_INTL
 * 
 *  Демонстрация использования метода WORKDAY_INTL класса ApiWorksheetFunction
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
        // This example shows how to return the serial number of the date before or after a specified number of workdays with custom weekend parameters.
        
        // How to return the serial number of the date adding some workdays.
        
        // Use a function to calculate the serial number of the date.
        
        const worksheet = Api.GetActiveSheet();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.WORKDAY_INTL("9/8/2017", "-20", "0000011", "8/15/2017");
        
        worksheet.GetRange("C1").SetValue(ans);
        
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
