/**
 * OnlyOffice JavaScript макрос - ApiParaPr.SetTabs
 * 
 *  Демонстрация использования метода SetTabs класса ApiParaPr
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
        // This example sets a sequence of custom tab stops which will be used for any tab characters in the paragraph.
        
        // How to change sizes of tabs between paragraphs.
        
        // Customize all kind of tabs indicating sizes.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 150 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        paraPr.SetTabs([1440, 2880, 4320], ["left", "center", "right"]);
        paragraph.AddTabStop();
        paragraph.AddText("Custom tab - 1 inch left");
        paragraph.AddLineBreak();
        paragraph.AddTabStop();
        paragraph.AddTabStop();
        paragraph.AddText("Custom tab - 2 inches center");
        paragraph.AddLineBreak();
        paragraph.AddTabStop();
        paragraph.AddTabStop();
        paragraph.AddTabStop();
        paragraph.AddText("Custom tab - 3 inches right");
        
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
