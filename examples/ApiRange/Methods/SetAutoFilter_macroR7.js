/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/SetAutoFilter.js
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
        // This example sets the autofilter by cell range.
        
        // How to automatically filter the specified range values.
        
        // Automatically filter out a range values.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("header");
        worksheet.GetRange("A2").SetValue("value2");
        worksheet.GetRange("A3").SetValue("value3");
        worksheet.GetRange("A4").SetValue("value4");
        worksheet.GetRange("A5").SetValue("value5");
        let range = worksheet.GetRange("A1:A5");
        range.SetAutoFilter(1, ["value2","value3"], "xlFilterValues");
        
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
