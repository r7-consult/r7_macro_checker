/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.POISSON_DIST
 * 
 *  Демонстрация использования метода POISSON_DIST класса ApiWorksheetFunction
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
        // This example shows how to calculate the Poisson distribution.
        
        // How to return the Poisson distribution.
        
        // Use a function to calculate the Poisson distribution.
        
        const worksheet = Api.GetActiveSheet();
        
        //method params
        let x = 9;
        let mean = 12;
        let cumulative = false;
        
        let func = Api.GetWorksheetFunction();
        let ans = func.POISSON_DIST(x, mean, cumulative);
        
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
