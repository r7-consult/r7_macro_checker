/**
 * OnlyOffice JavaScript макрос - ApiParagraph.SetIndLeft
 * 
 *  Демонстрация использования метода SetIndLeft класса ApiParagraph
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
        // This example sets the paragraph left side indentation.
        
        // How to change the sentence's left side indentation of the paragraph.
        
        // Get a paragraph from the shape's content then add a text specifying the left side indentation.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("This is a paragraph with the indent of 2 inches set to it. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.SetIndLeft(2880);
        paragraph = Api.CreateParagraph();
        paragraph.AddText("This is a paragraph without any indent set to it. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        content.Push(paragraph);
        
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
