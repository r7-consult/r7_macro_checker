/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/GetFormulaArray.js
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
        // This example shows how to get the array formula of a range.
        
        // How to get an array formula value.
        
        // Get a range, get its array formula value and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1:A3").SetFormulaArray("={1;2;3}");
        let value = worksheet.GetRange("A1").GetFormulaArray();
        worksheet.GetRange("A4").SetValue("Array formula of the cell A1: ");
        worksheet.GetRange("B4").SetValue("'" + value);
        value = worksheet.GetRange("A2").GetFormulaArray();
        worksheet.GetRange("A5").SetValue("Array formula of the cell A2: ");
        worksheet.GetRange("B5").SetValue("'" + value);
        value = worksheet.GetRange("A3").GetFormulaArray();
        worksheet.GetRange("A6").SetValue("Array formula of the cell A3: ");
        worksheet.GetRange("B6").SetValue("'" + value);
        
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
