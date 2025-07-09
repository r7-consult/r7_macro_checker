/**
 * OnlyOffice JavaScript макрос - ApiAreas.GetParent
 * 
 *  Демонстрация использования метода GetParent класса ApiAreas
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
        // This example shows how to get the parent object for the specified collection.
        
        // How to get a parent of the collection.
        
        // Find a collection parent of the selected range.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1:D1");
        range.SetValue("1");
        range.Select();
        let areas = range.GetAreas();
        let parent = areas.GetParent();
        let type = parent.GetClassType();
        range = worksheet.GetRange('A4');
        range.SetValue("The areas parent: ");
        range.AutoFit(false, true);
        worksheet.GetRange('B4').Paste(parent);
        range = worksheet.GetRange('A5');
        range.SetValue("The type of the areas parent: ");
        range.AutoFit(false, true);
        worksheet.GetRange('B5').SetValue(type);
        
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
