/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: ApiProtectedRange/Methods/GetUser.js
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
        // This example gets a user of a protected range.
        
        // How to get a user information of the protected range.
        
        // Get an active sheet, add protected range to it, add user with rights and get user info. 
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddProtectedRange("protectedRange", "$A$1:$B$1").AddUser("userId", "name", "CanView");
        let protectedRange = worksheet.GetProtectedRange("protectedRange");
        let userInfo = protectedRange.GetUser("userId");
        let userName = userInfo.GetName();
        worksheet.GetRange("A3").SetValue("User name: " + userName);
        
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
