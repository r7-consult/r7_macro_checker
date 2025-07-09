/**
 * OnlyOffice JavaScript макрос - ApiFont.SetSubscript
 * 
 *  Демонстрация использования метода SetSubscript класса ApiFont
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
        // This example sets the subscript property to the specified font.
        
        // How to change a regular text to a subscript.
        
        // Get a font object of characters and make it subscript.
        
        const worksheet = Api.GetActiveSheet();
        const range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        const characters = range.GetCharacters(9, 4);
        const font = characters.GetFont();
        font.SetSubscript(true);
        
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
