/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.PERCENTILE_EXC
 * 
 *  Демонстрация использования метода PERCENTILE_EXC класса ApiWorksheetFunction
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
        // This example shows how to return the k-th percentile of values in a range, where k is in the range 0..1, exclusive.
        
        // How to get the k-th percentile of values in a range (exclusive).
        
        // Use a function to get the k-th percentile of values.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let column1 = [1, 0, 7, 10];
        let column2 = [3, 2, 5, 8];
        let column3 = [5, 4, 3, 6];
        let column4 = [7, 6, 5, 4];
        
        for (let i = 0; i < column1.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(column1[i]);
        }
        for (let j = 0; j < column2.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(column2[j]);
        }
        for (let n = 0; n < column3.length; n++) {
            worksheet.GetRange("C" + (n + 1)).SetValue(column3[n]);
        }
        for (let m = 0; m < column4.length; m++) {
            worksheet.GetRange("D" + (m + 1)).SetValue(column4[m]);
        }
        
        let range = worksheet.GetRange("A1:D4");
        worksheet.GetRange("D5").SetValue(func.PERCENTILE_EXC(range, 0.5));
        
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
