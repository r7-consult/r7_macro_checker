/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiParaPr/Methods/GetClassType.js
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
        // This example gets a class type and inserts it into the document.
        
        // How to get a class type of ApiParaPr.
        
        // Get a class type of ApiParaPr and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        let classType = paraPr.GetClassType();
        paraPr.SetIndFirstLine(1440);
        paragraph.AddText("This is the first paragraph with the indent of 1 inch set to the first line. ");
        paragraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes.");
        paragraph = Api.CreateParagraph();
        paragraph.AddText("Class Type = " + classType);
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
