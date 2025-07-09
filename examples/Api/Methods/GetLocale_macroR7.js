/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/GetLocale.js
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
        // This example shows how to get the current locale ID.
        
        // How to set and get current locale ID.
        
        // Get region ID and insert information into the cell.
        
        let worksheet = Api.GetActiveSheet();Api.SetLocale("en-CA");
        let locale = Api.GetLocale();
        worksheet.GetRange("A1").SetValue("Locale: " + locale);
        
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
