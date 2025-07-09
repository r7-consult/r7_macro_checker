/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/SHEETS.js
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
        // This example shows how to return the number of sheets in a reference.
        
        // How to count sheets.
        
        // Use a function to get how many sheets are present in the worksheet.
        
        // Add more sheets
        
        Api.AddSheet("Sheet2")
        Api.AddSheet("Sheet3")
        
        // Get the number of sheets
        let func = Api.GetWorksheetFunction();
        let result = func.SHEETS();
        const worksheet = Api.GetActiveSheet();
        worksheet.GetRange("C3").SetValue(result);
        
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
