/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiParagraph/Methods/GetSpacingAfter.js
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
        // This example shows how to get the spacing after value of the paragraph.
        
        // How to get the spacing information which is after the paragraph.
        
        // Get two consecutive paragraphs, add the spacing between them then get the spacing after the first one and display it in the worksheet. 
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph1 = content.GetElement(0);
        paragraph1.AddText("This is an example of setting a space after a paragraph. ");
        paragraph1.AddText("The second paragraph will have an offset of one inch from the top. ");
        paragraph1.AddText("This is due to the fact that the first paragraph has this offset enabled.");
        paragraph1.SetSpacingAfter(1440);
        let paragraph2 = Api.CreateParagraph();
        paragraph2.AddText("This is the second paragraph and it is one inch away from the first paragraph.");
        paragraph2.AddLineBreak();
        let spacingAfter = paragraph1.GetSpacingAfter();
        paragraph2.AddText("Spacing after: " + spacingAfter);
        content.Push(paragraph2);
        
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
