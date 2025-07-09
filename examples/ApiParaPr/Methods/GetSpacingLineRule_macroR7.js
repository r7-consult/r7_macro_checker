/**
 * OnlyOffice JavaScript макрос - ApiParaPr.GetSpacingLineRule
 * 
 *  Демонстрация использования метода GetSpacingLineRule класса ApiParaPr
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
        // This example shows how to get the paragraph line spacing value.
        
        // How to get spacing line value between sentences of a paragraph.
        
        // Create a paragraph set spacing line between the sentences and retrieve the value.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        paraPr.SetSpacingLine(3 * 240, "auto");
        paragraph.AddText("Paragraph 1. Spacing: 3 times of a common paragraph line spacing.");
        paragraph.AddLineBreak();
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes.");
        let spacingLineValue = paraPr.GetSpacingLineValue();
        paragraph = Api.CreateParagraph();
        paragraph.AddText("Spacing line value : " + spacingLineValue);
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
