/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetAllPivotTables
 * 
 *  Демонстрация использования метода GetAllPivotTables класса ApiWorksheet
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
        // This example shows how to get all pivot tables from the sheet.
        
        // How to get all pivot tables.
        
        // Get all pivot tables as an array.
        
        let worksheet = Api.GetActiveSheet();
        
        worksheet.GetRange('B1').SetValue('Region');
        worksheet.GetRange('C1').SetValue('Price');
        worksheet.GetRange('B2').SetValue('East');
        worksheet.GetRange('B3').SetValue('West');
        worksheet.GetRange('C2').SetValue(42.5);
        worksheet.GetRange('C3').SetValue(35.2);
        
        let dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
        let pivotRef = worksheet.GetRange('A7');
        Api.InsertPivotExistingWorksheet(dataRef, worksheet.GetRange('A7'));
        Api.InsertPivotExistingWorksheet(dataRef, worksheet.GetRange('D7'));
        Api.InsertPivotExistingWorksheet(dataRef, worksheet.GetRange('G7'));
        
        worksheet.GetAllPivotTables().forEach(function (pivot) {
            pivot.AddDataField('Price');
        });
        
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
