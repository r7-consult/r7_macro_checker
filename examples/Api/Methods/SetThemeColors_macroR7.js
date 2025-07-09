/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/SetThemeColors.js
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
        // This example sets the theme colors to the current spreadsheet.
        
        // How to get all theme colors and apply one of them.
        
        // Apply one of the theme colors from the array of available ones.
        
        let worksheet = Api.GetActiveSheet();
        let themes = Api.GetThemesColors();
        for (let i = 0; i < themes.length; ++i) {
            worksheet.GetRange("A" + (i + 1)).SetValue(themes[i]);
        }
        Api.SetThemeColors(themes[3]);
        worksheet.GetRange("C3").SetValue("The 'Apex' theme colors were set to the current spreadsheet.");
        
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
