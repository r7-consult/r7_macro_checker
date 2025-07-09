/**
 * OnlyOffice JavaScript макрос - Api.RefreshAllPivots
 * 
 *  Демонстрация использования метода RefreshAllPivots класса Api
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
        // This example how to refresh all pivot tables in the active workbook.
        
        // How to refresh all pivot tables in a worksheet.
        
        // Refresh all values from the pivot table using a method.
        
        let worksheet = Api.GetActiveSheet();
        
        worksheet.GetRange('B1').SetValue('Region');
        worksheet.GetRange('C1').SetValue('Price');
        worksheet.GetRange('B2').SetValue('East');
        worksheet.GetRange('B3').SetValue('West');
        worksheet.GetRange('C2').SetValue(42.5);
        worksheet.GetRange('C3').SetValue(35.2);
        
        let dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
        let pivotTable = Api.InsertPivotNewWorksheet(dataRef);
        
        Api.GetPivotByName(pivotTable.GetName()).AddFields({
            rows: 'Region',
        });
        
        Api.GetPivotByName(pivotTable.GetName()).AddDataField('Price');
        Api.RefreshAllPivots();
        
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
