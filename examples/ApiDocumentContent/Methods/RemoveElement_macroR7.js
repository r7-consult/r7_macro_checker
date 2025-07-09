/**
 * OnlyOffice JavaScript макрос - ApiDocumentContent.RemoveElement
 * 
 *  Демонстрация использования метода RemoveElement класса ApiDocumentContent
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
        // This example removes an element using the position specified.
        
        // How to remove an element from a document knowing its position in the document content.
        
        // Delete an element from a document and prove it by showing the difference.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("This is paragraph #1.");
        for (let paraIncrease = 1; paraIncrease < 5; ++paraIncrease) {
            paragraph = Api.CreateParagraph();
            paragraph.AddText("This is paragraph #" + (paraIncrease + 1) + ".");
            content.Push(paragraph);
        }
        content.RemoveElement(2);
        paragraph = Api.CreateParagraph();
        paragraph.AddText("We removed paragraph #3, check that out above.");
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
