/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.GAMMA_INV
 * 
 *  Демонстрация использования метода GAMMA_INV класса ApiWorksheetFunction
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
        // This example shows how to return the inverse of the gamma cumulative distribution: if p = GAMMA.DIST(x,...), then GAMMA.INV(p,...) = x.
        
        // How to calculate the inverse of the gamma cumulative distribution.
        
        // Use a function to get the inverse of the gamma cumulative distribution.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.GAMMA_INV(0.4, 9, 2);
        worksheet.GetRange("B2").SetValue(ans);
        
        
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
