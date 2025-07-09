/**
 * Enhanced OnlyOffice JavaScript DSL Macro
 * Generated from: Api/Methods/GetFreezePanesType.js
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
        // This example freezes first column and get pastes a freezed type into the table.
        
        // How to freeze a column in a worksheet.
        
        // Freeze worksheet column and show its name in a cell.
        
        Api.SetFreezePanesType('column');
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("Type: ");
        worksheet.GetRange("B1").SetValue(Api.GetFreezePanesType());
        
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
