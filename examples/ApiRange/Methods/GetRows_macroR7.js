/**
 * OnlyOffice JavaScript макрос - ApiRange.GetRows
 * 
 *  Демонстрация использования метода GetRows класса ApiRange
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
        // This example shows how to get a Range object that represents the rows in the specified range.
        
        // How to get a cell rows of a range.
        
        // Get a range and change each cell's row value by getting all row objects.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("1:3");
        for (let i=1; i <= 3; i++) {
        	let rows = range.GetRows(i);    
        	rows.SetValue(i);
        }
        
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
