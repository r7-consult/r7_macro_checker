/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/onWorksheetChange.js
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
        // This example shows how to handle onWorksheetChange named event.
        
        // How to attach events, for example, onWorksheetChange one.
        
        // Event handling example in a worksheet.
        
        Api.attachEvent("onWorksheetChange", function(range){
            console.log("onWorksheetChange");
            console.log(range.GetAddress());
        });
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("1");
        
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
