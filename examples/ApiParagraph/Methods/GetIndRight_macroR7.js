/**
 * OnlyOffice JavaScript макрос - ApiParagraph.GetIndRight
 * 
 *  Демонстрация использования метода GetIndRight класса ApiParagraph
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
        // This example shows how to get the paragraph right side indentation.
        
        // How to get a right indent of a paragraph.
        
        // Get the right paragraph indent by the side.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("This is a paragraph with the right offset of 2 inches set to it. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.SetJc("right");
        paragraph.SetIndRight(2880);
        let indRight = paragraph.GetIndRight();
        paragraph = Api.CreateParagraph();
        paragraph.AddText("Right indent: " + indRight);
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
