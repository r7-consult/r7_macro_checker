/**
 * OnlyOffice JavaScript макрос - Api.GetCommentById
 * 
 *  Демонстрация использования метода GetCommentById класса Api
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
        // This example shows how to get a comment from the current document by its ID.
        
        // How to get specific comment by its ID.
        
        // Find a comment by its ID.
        
        let comment = Api.AddComment("Comment", "Bob");
        let id = comment.GetId();
        comment = Api.GetCommentById(id);
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("Comment Text: " + comment.GetText());
        worksheet.GetRange("B1").SetValue("Comment Author: " + comment.GetAuthorName());
        
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
