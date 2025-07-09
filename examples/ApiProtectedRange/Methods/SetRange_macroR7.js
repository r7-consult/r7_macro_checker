/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiProtectedRange/Methods/SetRange.js
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
        // This example changes ranges for a protected range.
        
        // How to set a protected range.
        
        // Update a range of a protected range.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddProtectedRange("protectedRange", "Sheet1!$A$1:$B$1");
        let protectedRange = worksheet.GetProtectedRange("protectedRange");
        protectedRange.SetRange("Sheet1!$A$2:$B$2");
        
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
