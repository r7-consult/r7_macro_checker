/**
 * OnlyOffice JavaScript макрос - ApiUniColor.GetClassType
 * 
 *  Демонстрация использования метода GetClassType класса ApiUniColor
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
        // This example gets a class type and pastes it into the presentation.
        
        // How to get a class type of ApiUniColor.
        
        // Get a class type of ApiUniColor and display it in the worksheet.
        
        const worksheet = Api.GetActiveSheet();
        const presetColor = Api.CreatePresetColor("peachPuff");
        const gs1 = Api.CreateGradientStop(presetColor, 0);
        const gs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
        const fill = Api.CreateLinearGradientFill([gs1, gs2], 5400000);
        const stroke = Api.CreateStroke(0, Api.CreateNoFill());
        worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 1, 3 * 36000);
        const classType = presetColor.GetClassType();
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
