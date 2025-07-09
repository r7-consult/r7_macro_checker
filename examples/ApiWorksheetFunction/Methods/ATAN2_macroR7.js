/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.ATAN2
 * 
 *  Демонстрация использования метода ATAN2 класса ApiWorksheetFunction
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
        // This example shows how to return the arctangent of the specified x and y coordinates, in radians between -Pi and Pi, excluding -Pi.
        
        // How to get an arctangent of the specified x and y coordinates.
        
        // Use function to get an arctangent of the specified x and y coordinates in radians.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ATAN2(1, -9));
        
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
