/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/GetCustomProperties.js
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
        // This example demonstrates how to use ApiCustomProperties to configure custom properties in a spreadsheet.
        
        const worksheet = Api.GetActiveSheet();
        
        const customProps = Api.GetCustomProperties();
        customProps.Add("MyStringProperty", "Hello, Spreadsheet!");
        customProps.Add("MyNumberProperty", 123.450);
        customProps.Add("MyDateProperty", new Date("23 November 2023"));
        customProps.Add("MyBoolProperty", true);
        
        let stringValue = customProps.Get("MyStringProperty");
        let numberValue = customProps.Get("MyNumberProperty");
        let dateValue = customProps.Get("MyDateProperty");
        let boolValue = customProps.Get("MyBoolProperty");
        
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(0, 100, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        const shape = worksheet.AddShape(
        	"rect",
        	100 * 36000, 50 * 36000,
        	fill, stroke,
        	0, 0, 5, 0
        );
        
        let docContent = shape.GetContent();
        let paragraph = docContent.GetElement(0);
        paragraph.AddText("Custom String Property: " + stringValue + "\n");
        paragraph.AddText("Custom Number Property: " + numberValue + "\n");
        paragraph.AddText("Custom Date Property: " + dateValue.toDateString() + "\n");
        paragraph.AddText("Custom Boolean Property: " + boolValue);
        
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
