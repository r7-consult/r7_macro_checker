/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetVisible
 * 
 *  Демонстрация использования метода GetVisible класса ApiWorksheet
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
        // This example shows how to get the state of sheet visibility.
        
        // How to get visibility of the worksheet.
        
        // Find out whether a sheet is visible or not and display it in the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetVisible(true);
        let isVisible = worksheet.GetVisible();
        worksheet.GetRange("A1").SetValue("Visible: ");
        worksheet.GetRange("B1").SetValue(isVisible);
        
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
