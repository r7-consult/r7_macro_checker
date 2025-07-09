/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/AddCustomFunction.js
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
        // This example calculates custom function result.
        
        // How to add custom function library.
        
        // How to use custom function.
        
        // How to add cell values using custom function library.
        
        Api.AddCustomFunctionLibrary("LibraryName", function(){
            /**
             * Function that returns the argument
             * @customfunction
             * @param {any} first First argument.
             * @returns {any} second Second argument.
            */
            Api.AddCustomFunction(function ADD(first, second) {
                return first + second;
            });
        });
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange('A1').SetValue('=ADD(1,2)');
        
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
