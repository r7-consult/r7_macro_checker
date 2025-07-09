/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheet/Methods/SetPageOrientation.js
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
        // This example sets the page orientation.
        
        // How to change a page orientation.
        
        // Set a page orientation and display it in the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetPageOrientation("xlPortrait");
        let pageOrientation = worksheet.GetPageOrientation();
        worksheet.GetRange("A1").SetValue("Page orientation: ");
        worksheet.GetRange("C1").SetValue(pageOrientation);
        
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
