/**
 * OnlyOffice JavaScript макрос - ApiRun.GetFontNames
 * 
 *  Демонстрация использования метода GetFontNames класса ApiRun
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
        // This example shows how to get all font names from all elements inside the run.
        
        // How to get all font names from the ApiRun object elements.
        
        // Get all font names from a text run as an array and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.SetFontFamily("Comic Sans MS");
        run.AddText("This is a text run with the font family set to 'Comic Sans MS'.");
        paragraph.AddElement(run);
        paragraph.AddLineBreak();
        let fontNames = run.GetFontNames();
        paragraph = Api.CreateParagraph();
        paragraph.AddText("Run font names: ");
        paragraph.AddLineBreak();
        for (let i = 0; i < fontNames.length; i++) {
            paragraph.AddText(fontNames[i]);
            paragraph.AddLineBreak();
        }
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
