/**
 * OnlyOffice JavaScript макрос - ApiParagraph.GetPrevious
 * 
 *  Демонстрация использования метода GetPrevious класса ApiParagraph
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
        // This example shows how to get the previous paragraph.
        
        // How to get the previous paragraph from the current one.
        
        // Add two paragraphs into the shape content then get the first one using the GetPrevious method.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        content.RemoveAllElements();
        let paragraph1 = Api.CreateParagraph();
        paragraph1.AddText("This is the first paragraph.");
        content.Push(paragraph1);
        let paragraph2 = Api.CreateParagraph();
        paragraph2.AddText("This is the second paragraph.");
        content.Push(paragraph2);
        let previousParagraph = paragraph2.GetPrevious();
        previousParagraph.SetBold(true);
        
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
