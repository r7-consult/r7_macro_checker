/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiFont/Methods/SetUnderline.js
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
        // This example sets an underline of the type specified in the request to the font.
        
        // How to change a regular text to an underlined one.
        
        // Get a font object of characters and make it underlined.
        
        const worksheet = Api.GetActiveSheet();
        const range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        const characters = range.GetCharacters(9, 4);
        const font = characters.GetFont();
        font.SetUnderline("xlUnderlineStyleSingle");
        
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
