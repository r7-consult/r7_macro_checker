/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.LOGEST
 * 
 *  Демонстрация использования метода LOGEST класса ApiWorksheetFunction
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
        // This example shows how to return statistics that describe an exponential curve matching known data points.
        
        // How to get the statistics of exponential curve matching the data points.
        
        // Use a function to return the statistics of exponential curve matching the data points.
        
        const worksheet = Api.GetActiveSheet();
        
        //configure function parameters
        let yValues = [1500, 1230, 1700, 1000, 980, 1470, 1560, 1640, 1420, 1100];
        let xValues = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
        let constant = true;
        let stats = false;
        
        //set values in cells
        for (let i = 0; i < yValues.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(yValues[i]);
        }
        for (let i = 0; i < xValues.length; i++) {
          worksheet.GetRange("B" + (i + 1)).SetValue(xValues[i]);
        }
        
        //get x and y ranges
        let yRange = worksheet.GetRange("A1:A10");
        let xRange = worksheet.GetRange("B1:B10");
        
        let func = Api.GetWorksheetFunction();
        //invoke LOGEST method
        let ans = func.LOGEST(yRange, xRange, constant, stats);
        
        //print answer
        worksheet.GetRange("D1").SetValue(ans);
        
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
