/**
 * OnlyOffice JavaScript макрос - ApiBullet.GetClassType
 * 
 *  Демонстрация использования метода GetClassType класса ApiBullet
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
        // This example shows how to get a type of the ApiBullet class and insert it into the table.
        
        // How to get a class type of a bullet and show it in the sheet.
        
        // Get a class type of an object.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let docContent = shape.GetContent();
        let paragraph = docContent.GetElement(0);
        let bullet = Api.CreateNumbering("ArabicParenR", 1);
        paragraph.SetBullet(bullet);
        paragraph.AddText(" This is an example of the bulleted paragraph.");
        paragraph = Api.CreateParagraph();
        paragraph.SetBullet(bullet);
        paragraph.AddText(" This is an example of the bulleted paragraph.");
        docContent.Push(paragraph);
        let classType = bullet.GetClassType();
        paragraph = Api.CreateParagraph();
        paragraph.SetJc("left");
        paragraph.AddText("Class Type = " + classType);
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
