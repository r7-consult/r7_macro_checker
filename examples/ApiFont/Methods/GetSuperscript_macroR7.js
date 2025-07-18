/**
 * OnlyOffice JavaScript макрос - ApiFont.GetSuperscript
 * 
 *  Демонстрация использования метода GetSuperscript класса ApiFont
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
        // This example shows how to get the superscript property of the specified font.
        
        // How to determine a font superscript property.
        
        // Get a boolean value that represents whether a font has a superscript property or not and show the value in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        font.SetSuperscript(true);
        let isSuperscript = font.GetSuperscript();
        worksheet.GetRange("B3").SetValue("Superscript property: " + isSuperscript);
        
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
