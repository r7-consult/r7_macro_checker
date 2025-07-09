/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/GetFormula.js
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
        // This example shows how to get a formula of the specified range.
        
        // How to find out a range formula.
        
        // Get a range, get its cell formula and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue(1);
        worksheet.GetRange("C1").SetValue(2);
        let range = worksheet.GetRange("A1");
        range.SetValue("=SUM(B1:C1)");
        let formula = range.GetFormula();
        worksheet.GetRange("A3").SetValue("Formula from cell A1: " + formula);
        
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
