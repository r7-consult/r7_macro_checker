/**
 * OnlyOffice JavaScript макрос - ApiAreas.GetCount
 * 
 *  Демонстрация использования метода GetCount класса ApiAreas
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
        // This example shows how to get a value that represents the number of objects in the collection.
        
        // How to get collection objects count.
        
        // How to get array length.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1:D1");
        range.SetValue("1");
        range.Select();
        let areas = range.GetAreas();
        let count = areas.GetCount();
        range = worksheet.GetRange('A5');
        range.SetValue("The number of ranges in the areas: ");
        range.AutoFit(false, true);
        worksheet.GetRange('B5').SetValue(count);
        
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
