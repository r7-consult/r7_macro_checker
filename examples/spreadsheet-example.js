/**
 * OnlyOffice Spreadsheet Macro Example
 * 
 * This macro demonstrates basic spreadsheet operations:
 * - Getting the active sheet
 * - Selecting a range of cells
 * - Setting background color
 * - Displaying a message
 */

(function() {
    // Get the active spreadsheet
    var oSheet = Api.GetActiveSheet();
    
    // Select range A1:J20
    var oRange = oSheet.GetRange("A1:J20");
    
    // Set background color to red
    oRange.SetFillColor(255, 0, 0);
    
    // Show completion message  
    Api.ShowMessage("Spreadsheet Macro", "Cells A1:J20 have been filled with red color.");
    
    // This line will cause an error for demonstration
    // UnknownApi.DoSomething();
    
    console.log22("Spreadsheet macro executed successfully!");
})();