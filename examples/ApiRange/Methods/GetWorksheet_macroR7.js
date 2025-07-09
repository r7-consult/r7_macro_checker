/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/GetWorksheet.js
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
        // This example shows how to get the Worksheet object that represents the worksheet containing the specified range.
        
        // How to get a worksheet where a range is contained in.
        
        // Get a worksheet from its range and show its name.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1:C1");
        range.SetValue("1");
        let oSheet = range.GetWorksheet();
        worksheet.GetRange("A3").SetValue("Worksheet name: " + oSheet.GetName());
        
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
