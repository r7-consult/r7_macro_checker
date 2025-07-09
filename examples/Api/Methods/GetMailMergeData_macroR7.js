/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/GetMailMergeData.js
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
        // This example shows how to get the mail merge data.
        
        // Get mail merge data from the worksheet.
        
        // How to get mail merge information using index.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetColumnWidth(0, 20);
        worksheet.GetRange("A1").SetValue("Email address");
        worksheet.GetRange("B1").SetValue("Greeting");
        worksheet.GetRange("C1").SetValue("First name");
        worksheet.GetRange("D1").SetValue("Last name");
        worksheet.GetRange("A2").SetValue("user1@example.com");
        worksheet.GetRange("B2").SetValue("Dear");
        worksheet.GetRange("C2").SetValue("John");
        worksheet.GetRange("D2").SetValue("Smith");
        worksheet.GetRange("A3").SetValue("user2@example.com");
        worksheet.GetRange("B3").SetValue("Hello");
        worksheet.GetRange("C3").SetValue("Kate");
        worksheet.GetRange("D3").SetValue("Cage");
        let mailMergeData = Api.GetMailMergeData(0);
        worksheet.GetRange("A5").SetValue("Mail merge data: " + mailMergeData);
        
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
