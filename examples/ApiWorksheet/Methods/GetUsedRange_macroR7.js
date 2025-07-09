/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheet/Methods/GetUsedRange.js
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
        // This example shows how to get the ApiRange object that represents the used range on the specified worksheet.
        
        // How to get used ranges from the worksheet.
        
        // Get used ranges and fill it with color.
        
        let worksheet = Api.GetActiveSheet();
        let usedRange = worksheet.GetUsedRange();
        usedRange.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
