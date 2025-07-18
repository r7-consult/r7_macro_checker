/**
 * OnlyOffice JavaScript макрос - ApiRun.SetVertAlign
 * 
 *  Демонстрация использования метода SetVertAlign класса ApiRun
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
        // This example specifies the alignment which will be applied to the contents of the current run in relation to the default appearance of the text run.
        
        // How to set vertical alignment of a text object.
        
        // Create a text run object, specify its vertical alignment as "baseline", "subscript" or "superscript".
        
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
        run.SetVertAlign("subscript");
        run.AddText("This is a text run with the text aligned below the baseline vertically. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.SetVertAlign("baseline");
        run.AddText("This is a text run with the text aligned by the baseline vertically. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.SetVertAlign("superscript");
        run.AddText("This is a text run with the text aligned above the baseline vertically.");
        paragraph.AddElement(run);
        
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
