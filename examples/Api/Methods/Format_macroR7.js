/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/Format.js
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
        // This example shows how to get a class formatted according to the instructions contained in the format expression.
        
        // How to set a format for a cell or a range using a format expression.
        
        // Change a format of a range using an expression.
        
        let worksheet = Api.GetActiveSheet();
        let format = Api.Format("123456", "$#,##0");
        worksheet.GetRange("A1").SetValue(format);
        
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
