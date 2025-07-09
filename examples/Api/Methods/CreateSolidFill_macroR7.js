/**
 * OnlyOffice JavaScript макрос - Api.CreateSolidFill
 * 
 *  Демонстрация использования метода CreateSolidFill класса Api
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
        // This example creates a solid fill to apply to the object using a selected solid color as the object background.
        
        // Create a solid fill to set a background color.
        
        // How to create a solid color to fill a shape.
        
        let worksheet = Api.GetActiveSheet();
        let rgbColor = Api.CreateRGBColor(255, 111, 61);
        let fill = Api.CreateSolidFill(rgbColor);
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 1, 3 * 36000);
        
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
