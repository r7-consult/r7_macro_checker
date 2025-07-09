/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/TRIMMEAN.js
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
        // This example shows how to return the mean of the interior portion of a set of data values.
        
        // How to delete the mean of the data values interior portion.
        
        // Use a function to remove the mean from the interior portion of a set of data values.
        
        let worksheet = Api.GetActiveSheet();
        let argumentsArr = [1, 2, 3, 4];
        
        // Place the numbers in cells
        for (let i = 0; i < argumentsArr.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(argumentsArr[i]);
        }
        
        // Get values from the range
        let data = worksheet.GetRange("A1:A4");
        
        // Calculate the TRIMMEAN of the range A1:A6
        let func = Api.GetWorksheetFunction();
        let result = func.TRIMMEAN(data, 0.6);
        worksheet.GetRange("B1").SetValue(result);
        
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
