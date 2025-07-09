/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetActiveCell
 * 
 *  Демонстрация использования метода GetActiveCell класса ApiWorksheet
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
        // This example shows how to get an object that represents an active cell.
        
        // How to get selected active cell.
        
        // Get an active cell and insert data to it.
        
        let worksheet = Api.GetActiveSheet();
        let activeCell = worksheet.GetActiveCell();
        activeCell.SetValue("This sample text was placed in an active cell.");
        
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
