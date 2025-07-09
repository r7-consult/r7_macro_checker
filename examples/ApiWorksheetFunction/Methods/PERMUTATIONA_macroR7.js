/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/PERMUTATIONA.js
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
        // This example shows how to return the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects.
        
        // How to return the number of permutations for a given number of objects with duplicates.
        
        // Use a function to claculate the number of permutations including duplicates.
        
        const worksheet = Api.GetActiveSheet();
        
        //method params
        let number = 32;
        let number_chosen = 2;
        
        worksheet.GetRange("A1").SetValue(number);
        worksheet.GetRange("B1").SetValue(number_chosen);
        
        let func = Api.GetWorksheetFunction();
        let ans = func.PERMUTATIONA(number, number_chosen);
        
        worksheet.GetRange("C1").SetValue(ans);
        
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
