/**
 * OnlyOffice JavaScript макрос - ApiCommentReply.SetAuthorName
 * 
 *  Демонстрация использования метода SetAuthorName класса ApiCommentReply
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
        // This example sets the comment reply author's name.
        
        // How to add author's name to the reply.
        
        // Add a reply to the comment and set author name, then show author name in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        let reply = comment.GetReply();
        reply.SetAuthorName("Mark Potato");
        worksheet.GetRange("A3").SetValue("Comment's reply author: ");
        worksheet.GetRange("B3").SetValue(reply.GetAuthorName());
        
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
