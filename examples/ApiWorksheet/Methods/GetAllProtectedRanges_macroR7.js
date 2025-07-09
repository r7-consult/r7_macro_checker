/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheet/Methods/GetAllProtectedRanges.js
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
        // This example shows how to get an object that represents all protected ranges.
        
        // How to get all protected ranges.
        
        // Get all protected ranges as an array.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddProtectedRange("protectedRange1", "Sheet1!$A$1:$B$1");
        worksheet.AddProtectedRange("protectedRange2", "Sheet1!$A$2:$B$2");
        let protectedRanges = worksheet.GetAllProtectedRanges();
        protectedRanges[0].SetTitle("protectedRangeNew1");
        protectedRanges[1].SetTitle("protectedRangeNew2");
        
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
