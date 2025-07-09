/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/LINEST.js
 * 
 * This macro demonstrates proper OnlyOffice API usage with:
 * - Error handling
 * - Comprehensive comments
 * - Production-ready code structure
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
        // This example shows how to return statistics that describe a linear trend matching known data points, by fitting a straight line using the least squares method.
        
        // How to get statistics that describe a linear trend matching known data points.
        
        // Use a function to return statistics that describe a linear trend matching known data points using the least squares method.
        
        const worksheet = Api.GetActiveSheet();
        
        //configure function parameters
        let yValues = [1.5, 2, 3];
        let xValues = [2, 3.1, 3.5];
        let constant = true;
        let stats = true;
        
        //set values in cells
        for (let i = 0; i < yValues.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(yValues[i]);
        }
        for (let n = 0; n < xValues.length; n++) {
          worksheet.GetRange("B" + (n + 1)).SetValue(xValues[n]);
        }
        
        //get x and y ranges
        let yRange = worksheet.GetRange("A1:A3");
        let xRange = worksheet.GetRange("B1:B3");
        let func = Api.GetWorksheetFunction();
        
        //invoke LINEST method
        let ans = func.LINEST(yRange, xRange, constant, stats);
        
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
