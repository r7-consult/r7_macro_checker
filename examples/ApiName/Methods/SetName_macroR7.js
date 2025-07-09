/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiName/Methods/SetName.js
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
        // This example sets a string value representing the object name.
        
        // How to rename an object.
        
        // Set a new name for an object and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        Api.AddDefName("name", "Sheet1!$A$1:$B$1");
        let defName = Api.GetDefName("name");
        defName.SetName("new_name");
        let newDefName = Api.GetDefName("new_name");
        worksheet.GetRange("A3").SetValue("The new name of the range: " + newDefName.GetName());
        
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
