/**
 * OnlyOffice JavaScript макрос - Api.AddComment
 * 
 *  Демонстрация использования метода AddComment класса Api
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
        // This example adds a comment to the document.
        
        // How to add comments in a worksheet.
        
        // Insert a comment into a cell.
        
        Api.AddComment("Comment 1", "Bob");
        Api.AddComment("Comment 2");
        let comments = Api.GetComments();
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("Comment Text: " + comments[0].GetText());
        worksheet.GetRange("B1").SetValue("Comment Author: " + comments[0].GetAuthorName());
        
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
