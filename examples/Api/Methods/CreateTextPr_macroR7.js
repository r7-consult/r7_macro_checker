/**
 * OnlyOffice JavaScript макрос - Api.CreateTextPr
 * 
 *  Демонстрация использования метода CreateTextPr класса Api
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
        // This example creates the empty text properties.
        
        // How to set custom properties for an empty text.
        
        // Change a new text properties like font size, font style, etc.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 80 * 36000, 50 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let docContent = shape.GetContent();
        docContent.RemoveAllElements();
        let textPr = Api.CreateTextPr();
        textPr.SetFontSize(30);
        textPr.SetBold(true);
        let paragraph = Api.CreateParagraph();
        paragraph.SetJc("left");
        paragraph.AddText("This is a sample text with the font size set to 30 and the font weight set to bold.");
        paragraph.SetTextPr(textPr);
        docContent.Push(paragraph);
        
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
