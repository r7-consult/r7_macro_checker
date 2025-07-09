/**
 * OnlyOffice JavaScript макрос - Api.GetThemesColors
 * 
 *  Демонстрация использования метода GetThemesColors класса Api
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
        // This example shows how to get a list of all the available theme colors for the spreadsheet.
        
        // Get all theme colors from the worksheet.
        
        // List all available theme colors.
        
        let worksheet = Api.GetActiveSheet();
        let themes = Api.GetThemesColors();
        for (let i = 0; i < themes.length; ++i) {
            worksheet.GetRange("A" + (i + 1)).SetValue(themes[i]);
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
