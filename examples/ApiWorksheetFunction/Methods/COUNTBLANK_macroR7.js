/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.COUNTBLANK
 * 
 *  Демонстрация использования метода COUNTBLANK класса ApiWorksheetFunction
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
        // This example shows how to counts a number of empty cells in a specified range of cells.
        
        // How to find a number of empty cells.
        
        // Use function to get empty cells count.
        
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
        let ans = func.COUNTBLANK(worksheet.GetRange("A1:C3"));
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
