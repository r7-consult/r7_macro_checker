/**
 * OnlyOffice JavaScript макрос - Api.CreateColorByName
 * 
 *  Демонстрация использования метода CreateColorByName класса Api
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
        // This example creates a color selecting it from one of the available color presets.
        
        // How to use a color from the preset.
        
        // Find a color by name and use it to change font color
        
        let worksheet = Api.GetActiveSheet();
        let color = Api.CreateColorByName("peachPuff");
        worksheet.GetRange("A2").SetValue("Text with color");
        worksheet.GetRange("A2").SetFontColor(color);
        
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
