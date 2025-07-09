/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/Delete.js
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
        // This example deletes the Range object.
        
        // How to remove a range from the worksheet.
        
        // Get a range from the worksheet and delete it specifying the direction.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B4").SetValue("1");
        worksheet.GetRange("C4").SetValue("2");
        worksheet.GetRange("D4").SetValue("3");
        worksheet.GetRange("C5").SetValue("5");
        let range = worksheet.GetRange("C4");
        range.Delete("up");
        
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
