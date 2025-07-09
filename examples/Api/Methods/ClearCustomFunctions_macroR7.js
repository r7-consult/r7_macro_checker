/**
 * OnlyOffice JavaScript макрос - Api.ClearCustomFunctions
 * 
 *  Демонстрация использования метода ClearCustomFunctions класса Api
 * https://r7-consult.ru/
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
        // This example clears all added custom functions.
        
        // How to delete custom function library.
        
        // Remove unused custom function library.
        
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
        worksheet.GetRange("A1").SetValue("=ADD(1, 2)");
        Api.ClearCustomFunctions();
        worksheet.GetRange("A3").SetValue("All the custom functions were removed.");
        
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
