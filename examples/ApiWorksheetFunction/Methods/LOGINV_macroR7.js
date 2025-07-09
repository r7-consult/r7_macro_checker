/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/LOGINV.js
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
        // This example shows how to return the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with the specified parameters.
        
        // How to get the lognormal cumulative distribution function of x.
        
        // Use a function to return the lognormal cumulative distribution function.
        
        const worksheet = Api.GetActiveSheet();
        
        //configure function parameters
        let numbersArr = [0.5, 4, 7];
        
        //set values in cells
        for (let i = 0; i < numbersArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(numbersArr[i]);
        }
        
        //get parameters
        let probability = worksheet.GetRange("A1");
        let mean = worksheet.GetRange("A2");
        let standardDeviation = worksheet.GetRange("A3");
        
        //invoke LOGINV method
        let func = Api.GetWorksheetFunction();
        let ans = func.LOGINV(probability, mean, standardDeviation);
        
        //print answer
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
