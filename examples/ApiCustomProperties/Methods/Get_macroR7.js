/**
 * OnlyOffice JavaScript макрос - ApiCustomProperties.Get
 * 
 *  Демонстрация использования метода Get класса ApiCustomProperties
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
        // This example demonstrates how to get the value of a custom property by its name.
        
        const worksheet = Api.GetActiveSheet();
        const customProps = Api.GetCustomProperties();
        
        customProps.Add("ExistingProp", "#123456");
        
        const existingProp = customProps.Get("ExistingProp");
        const nonExistentProp = customProps.Get("NonExistentProp");
        
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(0, 100, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        const shape = worksheet.AddShape(
        	"rect",
        	100 * 36000, 50 * 36000,
        	fill, stroke,
        	0, 0, 5, 0
        );
        
        let paragraph = shape.GetDocContent().GetElement(0);
        paragraph.AddText("Existing Property Value: " + existingProp);
        paragraph.AddText("\nNon-Existent Property Value: " + nonExistentProp);
        
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
