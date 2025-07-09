/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/GetAllComments.js
 * 
 * This macro demonstrates proper OnlyOffice API usage with:
 * - Error handling
 * - Comprehensive comments
 * - Production-ready code structure
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
        // This example shows how to get all comments from the worksheet.
        
        // How to get all comments from the worksheet.
        
        // Get all cell comments.
        
        let worksheet = Api.GetActiveSheet();Api.AddComment("Comment 1", "John Smith");
        worksheet.GetRange("A4").AddComment("Comment 2", "Mark Potato");
        let arrComments = Api.GetAllComments();
        worksheet.GetRange("A1").SetValue("Comment text: " + arrComments[1].GetText());
        worksheet.GetRange("A2").SetValue("Comment author: " + arrComments[1].GetAuthorName());
        
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
