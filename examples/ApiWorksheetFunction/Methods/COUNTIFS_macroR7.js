/**
 * OnlyOffice JavaScript макрос - ApiWorksheetFunction.COUNTIFS
 * 
 *  Демонстрация использования метода COUNTIFS класса ApiWorksheetFunction
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
        // This example shows how to count a number of cells specified by a given set of conditions or criteria.
        
        // How to find a number of cells that satisfy a list of conditions.
        
        // Use function to get cells if conditions are met.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let buyer = ["Buyer", "Tom", "Bob", "Ann", "Kate", "John"];
        let product = ["Product", "Apples", "Red apples", "ranges", "Green apples", "ranges"];
        let quantity = ["Quantity", 12, 45, 18, 26, 10];
        
        for (let i = 0; i < buyer.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(buyer[i]);
        }
        for (let j = 0; j < product.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(product[j]);
        }
        for (let n = 0; n < quantity.length; n++) {
            worksheet.GetRange("C" + (n + 1)).SetValue(quantity[n]);
        }
        
        let range1 = worksheet.GetRange("B2:B6");
        let range2 = worksheet.GetRange("C2:C6");
        worksheet.GetRange("E6").SetValue(func.COUNTIFS(range1, "*apples", range2, "45"));
        
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
