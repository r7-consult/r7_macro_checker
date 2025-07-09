/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiFreezePanes/Methods/FreezeAt.js
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
        // This example freezes the specified range in top-and-left-most pane of the worksheet.
        
        // How to freeze a specified range of panes.
        
        // Get freeze panes and freeze the specified part.
        
        let worksheet = Api.GetActiveSheet();
        let freezePanes = worksheet.GetFreezePanes();
        let range = Api.GetRange('H2:K4');
        freezePanes.FreezeAt(range);
        
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
