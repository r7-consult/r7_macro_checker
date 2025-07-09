/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/Replace.js
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
        // This example replaces specific information to another one in a range.
        
        // How to replace one data value with another in a range.
        
        // Create a range and replace its data field value with a new one.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue(2014);
        worksheet.GetRange("C1").SetValue(2015);
        worksheet.GetRange("D1").SetValue(2016);
        worksheet.GetRange("A2").SetValue("Projected Revenue");
        worksheet.GetRange("A3").SetValue("Estimated Costs");
        worksheet.GetRange("A4").SetValue("Cost price");
        worksheet.GetRange("B2").SetValue(200);
        worksheet.GetRange("B3").SetValue(250);
        worksheet.GetRange("B4").SetValue(50);
        worksheet.GetRange("C2").SetValue(200);
        worksheet.GetRange("C3").SetValue(260);
        worksheet.GetRange("C4").SetValue(120);
        worksheet.GetRange("D2").SetValue(200);
        worksheet.GetRange("D3").SetValue(200);
        worksheet.GetRange("D4").SetValue(160);
        let range = worksheet.GetRange("A2:D4");
        let replaceData = {
            What: "200", 
            Replacement: "0",
            LookAt: "xlWhole",
            SearchOrder: "xlByColumns",
            SearchDirection: "xlNext",
            MatchCase: true,
            ReplaceAll: true
        };
        range.Replace(replaceData);
        
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
