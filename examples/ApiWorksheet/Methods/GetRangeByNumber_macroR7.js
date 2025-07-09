/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheet/Methods/GetRangeByNumber.js
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
        // This example shows how to get an object that represents the selected range of the sheet using the row/column coordinates for the cell selection.
        
        // How to get a range using its coordinates.
        
        // Get range by number and set its value.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRangeByNumber(1, 2).SetValue("42");
        
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
