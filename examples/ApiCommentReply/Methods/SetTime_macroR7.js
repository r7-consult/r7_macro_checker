/**
 * OnlyOffice JavaScript макрос - ApiCommentReply.SetTime
 * 
 *  Демонстрация использования метода SetTime класса ApiCommentReply
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
        // This example sets the timestamp of the comment reply creation in the current time zone format.
        
        // How to change a time when a reply was created.
        
        // Add a reply then update its creation time and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        let reply = comment.GetReply();
        reply.SetTime(Date.now());
        worksheet.GetRange("A3").SetValue("Comment's reply timestamp: ");
        worksheet.GetRange("B3").SetValue(reply.GetTime());
        
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
