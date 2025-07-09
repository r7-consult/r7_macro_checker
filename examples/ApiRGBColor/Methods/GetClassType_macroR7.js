/**
 * OnlyOffice JavaScript макрос - ApiRGBColor.GetClassType
 * 
 *  Демонстрация использования метода GetClassType класса ApiRGBColor
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
        // This example gets a class type and inserts it into the document.
        
        // How to get a class type of ApiRGBColor.
        
        // Get a class type of ApiRGBColor and display it in the worksheet.
        
        const worksheet = Api.GetActiveSheet();
        const rgbColor = Api.CreateRGBColor(255, 213, 191);
        const gs1 = Api.CreateGradientStop(rgbColor, 0);
        const gs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
        const fill = Api.CreateLinearGradientFill([gs1, gs2], 5400000);
        const stroke = Api.CreateStroke(0, Api.CreateNoFill());
        worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 1, 3 * 36000);
        const classType = rgbColor.GetClassType();
        worksheet.SetColumnWidth(0, 15);
        worksheet.SetColumnWidth(1, 10);
        worksheet.GetRange("A1").SetValue("Class Type = ");
        worksheet.GetRange("B1").SetValue(classType);
        
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
