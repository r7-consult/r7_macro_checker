/**
 * OnlyOffice JavaScript макрос - ApiTextPr.GetSmallCaps
 * 
 *  Демонстрация использования метода GetSmallCaps класса ApiTextPr
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
        // This example gets a text capitalization using its property.
        
        // How to find out whether a text is uncapitalized or not.
        
        // Find whether a text characters are in small caps or not.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text. ");
        run.AddText("The text properties are changed and the style is added to the paragraph. ");
        run.AddLineBreak();
        paragraph.AddElement(run);
        let textProps = run.GetTextPr();
        textProps.SetSmallCaps(true);
        paragraph = Api.CreateParagraph();
        let isSmallCaps = textProps.GetSmallCaps();
        paragraph.AddText("Property of the small capitalized letters: " + isSmallCaps);
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
