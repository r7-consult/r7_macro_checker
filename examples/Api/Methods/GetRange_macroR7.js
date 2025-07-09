/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/GetRange.js
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
        // This example shows how to get the ApiRange object by the range reference.
        
        // How to get a range object using an address.
        
        // Update a rabge knowing its cell address.
        
        let worksheet = Api.GetActiveSheet();
        let range = Api.GetRange("A1:C1");
        range.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        worksheet.GetRange("A3").SetValue("The color was set to the background of cells A1:C1.");
        
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
