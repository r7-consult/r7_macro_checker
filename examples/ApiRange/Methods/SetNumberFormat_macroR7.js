/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/SetNumberFormat.js
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
        // This example specifies whether a number in the cell should be treated like number, currency, date, time, etc. or just like text.
        
        // How to set number format of cells.
        
        // Get a range and specify number format of its cells.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetColumnWidth(0, 30);
        worksheet.SetColumnWidth(1, 30);
        worksheet.GetRange("A2").SetNumberFormat("General");
        worksheet.GetRange("A2").SetValue("123456");
        worksheet.GetRange("B2").SetValue("General");
        worksheet.GetRange("A3").SetNumberFormat("0.00");
        worksheet.GetRange("A3").SetValue("123456");
        worksheet.GetRange("B3").SetValue("Number");
        worksheet.GetRange("A4").SetNumberFormat("$#,##0.00");
        worksheet.GetRange("A4").SetValue("123456");
        worksheet.GetRange("B4").SetValue("Currency");
        worksheet.GetRange("A5").SetNumberFormat("_($* #,##0.00_)");
        worksheet.GetRange("A5").SetValue("123456");
        worksheet.GetRange("B5").SetValue("Accounting");
        worksheet.GetRange("A6").SetNumberFormat("m/d/yyyy");
        worksheet.GetRange("A6").SetValue("123456");
        worksheet.GetRange("B6").SetValue("DateShort");
        worksheet.GetRange("A7").SetNumberFormat("[$-F800]dddd, mmmm dd, yyyy");
        worksheet.GetRange("A7").SetValue("123456");
        worksheet.GetRange("B7").SetValue("DateLong");
        worksheet.GetRange("A8").SetNumberFormat("[$-F400]h:mm:ss AM/PM");
        worksheet.GetRange("A8").SetValue("123456");
        worksheet.GetRange("B8").SetValue("Time");
        worksheet.GetRange("A9").SetNumberFormat("0.00%");
        worksheet.GetRange("A9").SetValue("123456");
        worksheet.GetRange("B9").SetValue("Percentage");
        worksheet.GetRange("A10").SetNumberFormat("0%");
        worksheet.GetRange("A10").SetValue("123456");
        worksheet.GetRange("B10").SetValue("Percent");
        worksheet.GetRange("A11").SetNumberFormat("# ?/?");
        worksheet.GetRange("A11").SetValue("123456");
        worksheet.GetRange("B11").SetValue("Fraction");
        worksheet.GetRange("A12").SetNumberFormat("0.00E+00");
        worksheet.GetRange("A12").SetValue("123456");
        worksheet.GetRange("B12").SetValue("Scientific");
        worksheet.GetRange("A13").SetNumberFormat("@");
        worksheet.GetRange("A13").SetValue("123456");
        worksheet.GetRange("B13").SetValue("Text");
        
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
