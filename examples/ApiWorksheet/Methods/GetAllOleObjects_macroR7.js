/**
 * OnlyOffice JavaScript макрос - ApiWorksheet.GetAllOleObjects
 * 
 *  Демонстрация использования метода GetAllOleObjects класса ApiWorksheet
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
        // This example shows how to get all OLE objects from the sheet.
        
        // How to get all OLE objects images.
        
        // Get all OLE objects as an array.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddOleObject("https://i.ytimg.com/vi_webp/SKGz4pmnpgY/sddefault.webp", 130 * 36000, 90 * 36000, "https://youtu.be/SKGz4pmnpgY", "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", 0, 2 * 36000, 4, 3 * 36000);
        let oleObjects = worksheet.GetAllOleObjects();
        let appId = oleObjects[0].GetApplicationId();
        worksheet.GetRange("A1").SetValue("The application ID for the current OLE object: " + appId);
        
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
