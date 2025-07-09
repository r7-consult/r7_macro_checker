/**
 * OnlyOffice JavaScript макрос - ApiComment.RemoveReplies
 * 
 *  Демонстрация использования метода RemoveReplies класса ApiComment
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
        // This example removes the specified comment replies.
        
        // How to remove replies from the comment.
        
        // Add a comment and replies to it, then remove specified reply and then show replies count.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        comment.AddReply("Reply 2", "John Smith", "uid-1");
        comment.RemoveReplies(0, 1, false);
        worksheet.GetRange("A3").SetValue("Comment replies count: ");
        worksheet.GetRange("B3").SetValue(comment.GetRepliesCount());
        
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
