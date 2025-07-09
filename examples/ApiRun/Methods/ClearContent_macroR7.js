/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiRun/Methods/ClearContent.js
 * 
 * This macro demonstrates proper OnlyOffice API usage with:
 * - Error handling
 * - Comprehensive comments
 * - Production-ready code structure
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
        // This example clears the content from the run.
        
        // How to create a text run object, add a text to it and clear its value.
        
        // Clear content of an ApiRun object.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.SetFontSize(30);
        run.AddText("This is just a sample text. ");
        run.AddText("But you will not see it in the resulting document, as it will be cleared.");
        paragraph.AddElement(run);
        run.ClearContent();
        paragraph = Api.CreateParagraph();
        run = Api.CreateRun();
        run.AddText("The text in the previous paragraph cannot be seen, as it has been cleared.");
        paragraph.AddElement(run);
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
