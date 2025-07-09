/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/TREND.js
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
        // This example shows how to return numbers in a linear trend matching known data points, using the least squares method.
        
        // How to get numbers in a linear trend using the least squares method.
        
        // Use a function to find a linear trend using data points by the least squares method.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let month = ["Month", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
        let sales = ["Sales", "$1,500.00", "$1,230.00", "$1,700.00", "$1,000.00", "$980.00", "$1,470.00", "$1,560.00", "$1,640.00", "$1,420.00", "$1,100.00"];
        
        for (let i = 0; i < month.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(month[i]);
        }
        for (let j = 0; j < sales.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(sales[j]);
        }
        
        worksheet.GetRange("C1").SetValue("Trend");
        let range1 = worksheet.GetRange("B2:B11");
        let range2 = worksheet.GetRange("A2:A11");
        worksheet.GetRange("C2:C11").SetValue(func.TREND(range1, range2));
        
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
