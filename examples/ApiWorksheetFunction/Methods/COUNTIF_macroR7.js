/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/COUNTIF.js
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
        // This example shows how to count a number of cells within a range that meet the given condition.
        
        // How to find a number of cells that satisfies some condition.
        
        // Use function to get cells if a condition is met.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let fruits = ["Apples", "ranges", "Bananas"];
        let numbers = [45, 6, 8];
        
        for (let i = 0; i < fruits.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(fruits[i]);
        }
        for (let j = 0; j < numbers.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(numbers[j]);
        }
        
        let range = worksheet.GetRange("A1:B3");
        worksheet.GetRange("D3").SetValue(func.COUNTIF(range, "*es"));
        
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
