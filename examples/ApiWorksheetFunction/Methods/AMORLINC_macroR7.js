/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.AMORLINC
 * 
 *  Демонстрация использования метода AMORLINC класса ApiWorksheetFunction
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
        // This example shows how to return the prorated linear depreciation of an asset for each accounting period.
        
        // How to get a prorated linear depreciation of an asset for each accounting period and display it in the worksheet.
        
        // Get a function that gets prorated linear depreciation of an asset for each accounting period.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.AMORLINC(3500, "1/1/2018", "3/1/2018", 500, 1, 0.25, 1));
        
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
