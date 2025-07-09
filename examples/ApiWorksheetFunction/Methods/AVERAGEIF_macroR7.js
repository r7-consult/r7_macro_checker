/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/AVERAGEIF.js
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
        // This example shows how to find the average (arithmetic mean) for the cells specified by a given condition or criteria.
        
        // How to find an average (arithmetic mean) using condition.
        
        // Use function to get an average of the cells if the condition is met.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let numbers = [67, 87, 98, 45];
        
        for (let i = 0; i < numbers.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(numbers[i]);
        }
        
        let range = worksheet.GetRange("A1:A4");
        worksheet.GetRange("C1").SetValue(func.AVERAGEIF(range, ">45"));
        
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
