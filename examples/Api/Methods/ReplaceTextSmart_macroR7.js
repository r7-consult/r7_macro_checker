/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/ReplaceTextSmart.js
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
        // This example replaces each paragraph (or text in cell) in the select with the corresponding text from an array of strings.
        
        // Replace string values of the selected range with a new values.
        
        // Replace cell string values with a new ones.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("A2").SetValue("2");
        let range = worksheet.GetRange("A1:A2");
        range.Select();
        Api.ReplaceTextSmart(["Cell 1", "Cell 2"]);
        
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
