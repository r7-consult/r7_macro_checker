/**
 * OnlyOffice JavaScript макрос - Api.GetComments
 * 
 *  Демонстрация использования метода GetComments класса Api
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
        // This example shows how to get an array of ApiComment objects.
        
        // How to get an array of comments.
        
        // Get all comments as an array.
        
        Api.AddComment("Comment 1", "Bob");
        Api.AddComment("Comment 2", "Bob");
        let arrComments = Api.GetComments();
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("Comment Text: " + arrComments[0].GetText());
        worksheet.GetRange("B1").SetValue("Comment Author: " + arrComments[0].GetAuthorName());
        
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
