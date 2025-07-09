/**
 * OnlyOffice JavaScript макрос - ApiComment.Delete
 * 
 *  Демонстрация использования метода Delete класса ApiComment
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
        // This example deletes the ApiComment object.
        
        // How to remove a comment from a range.
        
        // Get a range add a comment to it and then remove it.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        range.AddComment("This is just a number.");
        let comment = range.GetComment();
        comment.Delete();
        worksheet.GetRange("A3").SetValue("The comment was just deleted from A1.");
        
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
