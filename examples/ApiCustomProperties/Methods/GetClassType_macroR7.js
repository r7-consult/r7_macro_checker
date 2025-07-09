/**
 * OnlyOffice JavaScript макрос - ApiCustomProperties.GetClassType
 * 
 *  Демонстрация использования метода GetClassType класса ApiCustomProperties
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
        // This example demonstrates how to get the class type of ApiCustomProperties.
        
        const worksheet = Api.GetActiveSheet();
        const customProps = Api.GetCustomProperties();
        const classType = customProps.GetClassType();
        
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(0, 100, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        const shape = worksheet.AddShape("rect", 100 * 36000, 50 * 36000, fill, stroke, 0, 0, 5, 0);
        
        let paragraph = shape.GetDocContent().GetElement(0);
        paragraph.AddText("ApiCustomProperties class type: " + classType);
        
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
