/**
 * OnlyOffice JavaScript макрос - ApiFreezePanes.Unfreeze
 * 
 *  Демонстрация использования метода Unfreeze класса ApiFreezePanes
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
        // This example freezes first column then unfreeze all panes in the worksheet.
        
        // How to unfreeze columns from freezed panes.
        
        // Add freezed panes then unfreeze the first column and show all freezed ones' location to prove it.
        
        Api.SetFreezePanesType('column');
        let worksheet = Api.GetActiveSheet();
        let freezePanes = worksheet.GetFreezePanes();
        freezePanes.Unfreeze();
        let range = freezePanes.GetLocation();
        worksheet.GetRange("A1").SetValue("Location: ");
        worksheet.GetRange("B1").SetValue(range + "");
        
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
