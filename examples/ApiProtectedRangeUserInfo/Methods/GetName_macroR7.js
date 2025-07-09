/**
 * OnlyOffice JavaScript макрос - ApiProtectedRangeUserInfo.GetName
 * 
 *  Демонстрация использования метода GetName класса ApiProtectedRangeUserInfo
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
        // This example gets name of a protected range user.
        
        // How to get a user info of a protected range and show its name.
        
        // Get a user name of a protected range and add it to the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddProtectedRange("protectedRange", "$A$1:$B$1").AddUser("userId", "name", "CanView");
        let protectedRange = worksheet.GetProtectedRange("protectedRange");
        let userInfo = protectedRange.GetUser("userId");
        let userName = userInfo.GetName();
        worksheet.GetRange("A3").SetValue("Name: " + userName);
        
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
