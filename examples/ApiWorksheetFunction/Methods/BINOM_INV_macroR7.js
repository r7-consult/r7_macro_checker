/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.BINOM_INV
 * 
 *  Демонстрация использования метода BINOM_INV класса ApiWorksheetFunction
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
        // This example shows how to return the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value. 
        
        // How to get a smallest value for which the cumulative binomial distribution >= criterion value.
        
        // Use function to get a minimum value so that the cumulative binomial distribution >= criterion value.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.BINOM_INV(678, 0.1, 0.007);
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
