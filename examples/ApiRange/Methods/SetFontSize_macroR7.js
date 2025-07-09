/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/SetFontSize.js
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
        // This example sets the font size to the characters of the cell range.
        
        // How to resize a cell font size.
        
        // Get a range and set its font size.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("2");
        let range = worksheet.GetRange("A1:D5");
        range.SetFontSize(20);
        
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
