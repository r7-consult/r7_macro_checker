/**
 * OnlyOffice JavaScript макрос - ApiRange.AddComment
 * 
 *  Демонстрация использования метода AddComment класса ApiRange
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
        // This example adds a comment to the range.
        
        // How to comment a range.
        
        // Get a range from the worksheet, add a comment to it and then show the comments text.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("1");
        range.AddComment("This is just a number.");
        worksheet.GetRange("A3").SetValue("The comment was added to the cell A1.");
        worksheet.GetRange("A4").SetValue("Comment: " + range.GetComment().GetText());
        
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
