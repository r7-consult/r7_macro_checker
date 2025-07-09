/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiAreas/Methods/GetItem.js
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
        // This example shows how to get a single object from a collection by its ID.
        
        // How to find an object by its ID from the collection.
        
        // Get element from an array by its ID.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1:D1");
        range.SetValue("1");
        range.Select();
        let areas = range.GetAreas();
        let item = areas.GetItem(1);
        range = worksheet.GetRange('A5');
        range.SetValue("The first item from the areas: ");
        range.AutoFit(false, true);
        worksheet.GetRange('B5').Paste(item);
        
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
