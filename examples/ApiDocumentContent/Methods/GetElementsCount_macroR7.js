/**
 * OnlyOffice JavaScript макрос - ApiDocumentContent.GetElementsCount
 * 
 *  Демонстрация использования метода GetElementsCount класса ApiDocumentContent
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
        // This example shows how to get a number of elements in the current document content.
        
        // How to get a number of elements of a shape from a document content.
        
        // Get a shape than count number of elements and display it using paragraph.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("We got the first paragraph inside the shape.");
        paragraph.AddLineBreak();
        paragraph.AddText("Number of elements inside the shape: " + content.GetElementsCount());
        paragraph.AddLineBreak();
        paragraph.AddText("Line breaks are NOT counted into the number of elements.");
        
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
