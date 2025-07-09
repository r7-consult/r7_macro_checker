/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiWorksheetFunction/Methods/FORECAST_ETS_SEASONALITY.js
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
        // This example shows how to return the length of the repetitive pattern an application detects for the specified time series.
        
        // How to return the length of the repetitive pattern.
        
        // Use function to return the length of the repetitive pattern for the specified time series.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let dates = ["10/1/2017", "11/1/2017", "12/1/2017", "1/1/2018", "2/1/2018"];
        let numbers = [12558, 14356, 16345, 18678, 14227];
        
        for (let i = 0; i < dates.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(dates[i]);
        }
        for (let j = 0; j < numbers.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(numbers[j]);
        }
        
        let range1 = worksheet.GetRange("B1:B5");
        let range2 = worksheet.GetRange("A1:A5");
        worksheet.GetRange("B6").SetValue(func.FORECAST_ETS_SEASONALITY(range1, range2, 1, 1));
        
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
