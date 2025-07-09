/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/MAX.js
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
        // This example shows how to return the largest value in a set of values. Ignores logical values and text.
        
        // How to get a maximum number from a list of numbers.
        
        // Use a function to find a maximum from a list.
        
        const worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let max = func.MAX(123, 197, 46, 354, 67, 456);
        worksheet.GetRange("C1").SetValue(max);
        
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
