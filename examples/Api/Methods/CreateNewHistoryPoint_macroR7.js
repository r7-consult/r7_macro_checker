/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/CreateNewHistoryPoint.js
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
        // This example creates a new history point.
        
        // Add history point for a range.
        
        // How to make a history point.
        
        var worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("This is just a sample text.");
        Api.CreateNewHistoryPoint();
        worksheet.GetRange("A3").SetValue("New history point was just created.");
        
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
