/**
 * OnlyOffice JavaScript макрос - ApiProtectedRange.DeleteUser
 * 
 *  Демонстрация использования метода DeleteUser класса ApiProtectedRange
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
        // This example deletes the the user protected range.
        
        // How to close an access for the protected range to user specifing user id, name and access type.
        
        // Get an active sheet, add protected range to it, add users with rights then delete one of them. 
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddProtectedRange("protectedRange", "$A$1:$B$1");
        let protectedRange = worksheet.GetProtectedRange("protectedRange");
        protectedRange.AddUser("userId", "name", "CanView");
        protectedRange.DeleteUser("userId");
        
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
