/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRange/Methods/SetUnderline.js
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
        // This example specifies that the contents of the current cell is displayed along with a line appearing directly below the character.
        
        // How to add underline to the cell value.
        
        // Get a range and add underline to its text.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("The text underlined with a single line");
        worksheet.GetRange("A2").SetUnderline("single");
        worksheet.GetRange("A4").SetValue("Normal text");
        
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
