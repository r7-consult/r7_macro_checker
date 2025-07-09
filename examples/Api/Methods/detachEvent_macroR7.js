/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/detachEvent.js
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
        // This example unsubscribes from the "onWorksheetChange" event.
        
        // Detach from an event.
        
        // How to stop event handling.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("1");
        Api.attachEvent("onWorksheetChange", function(range){
            console.log("onWorksheetChange");
            console.log(range.GetAddress());
        });
        Api.detachEvent("onWorksheetChange");
        
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
