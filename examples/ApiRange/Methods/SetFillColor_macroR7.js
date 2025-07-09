/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/SetFillColor.js
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
        // This example sets the background color to the cell range with the previously created color object.
        
        // How to color a cell.
        
        // Get a range and apply a solid fill to its background color.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetColumnWidth(0, 50);
        worksheet.GetRange("A2").SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        worksheet.GetRange("A2").SetValue("This is the cell with a color set to its background");
        worksheet.GetRange("A4").SetValue("This is the cell with a default background color");
        
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
