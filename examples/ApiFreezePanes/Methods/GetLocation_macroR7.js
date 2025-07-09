/**
 * OnlyOffice JavaScript макрос - ApiFreezePanes.GetLocation
 * 
 *  Демонстрация использования метода GetLocation класса ApiFreezePanes
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
        // This example freezes first column and get pastes a freezed range address into the table.
        
        // How to get location address of a freezed column.
        
        // Get an address of a column from freezed panes and display it in the worksheet.
        
        Api.SetFreezePanesType('column');
        let worksheet = Api.GetActiveSheet();
        let freezePanes = worksheet.GetFreezePanes();
        let range = freezePanes.GetLocation();
        worksheet.GetRange("A1").SetValue("Location: ");
        worksheet.GetRange("B1").SetValue(range.GetAddress());
        
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
