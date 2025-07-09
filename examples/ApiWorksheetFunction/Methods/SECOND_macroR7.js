/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.SECOND
 * 
 *  Демонстрация использования метода SECOND класса ApiWorksheetFunction
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
        // This example shows how to return the second, a number from 0 to 59.
        
        // How to get seconds from time.
        
        // Use a function to return seconds from time object.
        
        const worksheet = Api.GetActiveSheet();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.SECOND("16:39:35"); 
        
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
