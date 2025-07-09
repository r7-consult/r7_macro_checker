/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.QUARTILE_INC
 * 
 *  Демонстрация использования метода QUARTILE_INC класса ApiWorksheetFunction
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
        // This example shows how to return the quartile of a data set, based on percentile values from 0..1, inclusive.
        
        // How to get the inclusive quartile of a data set.
        
        // Use a function to calculate an inclusive fourth part of a data set.
        
        const worksheet = Api.GetActiveSheet();
        
        let valueArr1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
        let quart = 1;
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr1.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr1[i]);
        }
        
        //method params
        let array = worksheet.GetRange("A1:A12");
        
        let func = Api.GetWorksheetFunction();
        let ans = func.QUARTILE_INC(array, quart); //0...1 inclusive
        
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
