/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/GetDocumentInfo.js
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
        // This example shows how to get the document info represented as an object and paste the application name into "A1" cell.
        
        // How to get document info and iys application name.
        
        // Get application name using document info.
        
        let docInfo = Api.GetDocumentInfo();
        let range = Api.GetActiveSheet().GetRange('A1');
        range.SetValue('This document has been created with: ' + docInfo.Application);
        
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
