/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiPivotDataField/Methods/GetPosition.js
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
        // This example shows how to get a position of a data field.
        
        // How to get a table element position.
        
        // Create a pivot table, add data to it then display a custom data field's position.
        
        let worksheet = Api.GetActiveSheet();
        
        worksheet.GetRange('B1').SetValue('Region');
        worksheet.GetRange('C1').SetValue('Style');
        worksheet.GetRange('D1').SetValue('Price');
        
        worksheet.GetRange('B2').SetValue('East');
        worksheet.GetRange('B3').SetValue('West');
        worksheet.GetRange('B4').SetValue('East');
        worksheet.GetRange('B5').SetValue('West');
        
        worksheet.GetRange('C2').SetValue('Fancy');
        worksheet.GetRange('C3').SetValue('Fancy');
        worksheet.GetRange('C4').SetValue('Tee');
        worksheet.GetRange('C5').SetValue('Tee');
        
        worksheet.GetRange('D2').SetValue(42.5);
        worksheet.GetRange('D3').SetValue(35.2);
        worksheet.GetRange('D4').SetValue(12.3);
        worksheet.GetRange('D5').SetValue(24.8);
        
        let dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
        let pivotTable = Api.InsertPivotNewWorksheet(dataRef);
        
        pivotTable.AddFields({
            rows: ['Region', 'Style'],
        });
        
        pivotTable.AddDataField('Price');
        let dataField = pivotTable.AddDataField('Price');
        
        let pivotWorksheet = Api.GetActiveSheet();
        
        pivotWorksheet.GetRange('A15').SetValue('Sum of Price2 position:');
        pivotWorksheet.GetRange('B15').SetValue(dataField.GetPosition());
        
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
