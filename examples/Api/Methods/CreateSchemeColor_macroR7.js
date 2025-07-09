/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/CreateSchemeColor.js
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
        // This example creates a complex color scheme selecting from one of the available schemes.
        
        // Get a color scheme using its name.
        
        // How to create a color from the schemes.
        
        let worksheet = Api.GetActiveSheet();
        let schemeColor = Api.CreateSchemeColor("dk1");
        let fill = Api.CreateSolidFill(schemeColor);
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        worksheet.AddShape("curvedUpArrow", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 1, 3 * 36000);
        
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
