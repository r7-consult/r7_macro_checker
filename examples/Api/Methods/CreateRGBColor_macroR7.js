/**
 * OnlyOffice JavaScript макрос - Api.CreateRGBColor
 * 
 *  Демонстрация использования метода CreateRGBColor класса Api
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
        // This example creates an RGB color setting the appropriate values for the red, green and blue color components.
        
        // How to create an RGB color.
        
        // Use RGB (R-red, G-green, B-blue) values to create colors.
        
        let worksheet = Api.GetActiveSheet();
        let gs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
        let gs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
        let fill = Api.CreateLinearGradientFill([gs1, gs2], 5400000);
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
