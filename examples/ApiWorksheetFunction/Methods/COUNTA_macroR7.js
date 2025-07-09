/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.COUNTA
 * 
 *  Демонстрация использования метода COUNTA класса ApiWorksheetFunction
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
        // This example shows how to count a number of cells in a range that are not empty.
        
        // How to find a number of non-empty cells.
        
        // Use function to get non-empty cells count.
        
        let worksheet = Api.GetActiveSheet();
        let numbersArr = [45, 6, 8];
        let stringsArr = ["Apples", "ranges", "Bananas"]
        
        // Place the numbers in cells
        for (let i = 0; i < numbersArr.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(numbersArr[i]);
        }
        
        // Place the strings in cells
        for (let n = 0; n < stringsArr.length; n++) {
            worksheet.GetRange("B" + (n + 1)).SetValue(stringsArr[n]);
        }
        
        let func = Api.GetWorksheetFunction();
        let ans = func.COUNTA(worksheet.GetRange("A1:C3"));
        worksheet.GetRange("D3").SetValue(ans);
        
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
