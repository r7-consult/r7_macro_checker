/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiColor/Methods/GetClassType.js
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
        // This example gets a class type and inserts it into the table.
        
        // How to get a color class type.
        
        // Get an RGB color type.
        
        let worksheet = Api.GetActiveSheet();
        let color = Api.CreateColorFromRGB(255, 111, 61);
        worksheet.GetRange("A2").SetValue("Text with color");
        worksheet.GetRange("A2").SetFontColor(color);
        let colorClassType = color.GetClassType();
        worksheet.GetRange("A4").SetValue("Class type = " + colorClassType);
        
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
