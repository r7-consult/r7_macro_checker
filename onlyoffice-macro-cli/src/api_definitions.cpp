#include "api_definitions.h"

namespace onlyoffice {
namespace macro {

void setupKnownAPIs(std::set<std::string>& knownAPIs) {
    // Main Api object
    knownAPIs.insert("Api");
    
    // Document & Worksheet Management
    knownAPIs.insert("Api.AddSheet");
    knownAPIs.insert("Api.GetActiveSheet");
    knownAPIs.insert("Api.GetActiveDocument");
    knownAPIs.insert("Api.GetActivePresentation");
    knownAPIs.insert("Api.CreateDocument");
    knownAPIs.insert("Api.GetWorkbook");
    knownAPIs.insert("Api.GetWorksheet");
    knownAPIs.insert("Api.GetSheet");
    knownAPIs.insert("Api.GetSheets");
    knownAPIs.insert("Api.Save");
    knownAPIs.insert("Api.ShowMessage");
    
    // Comments
    knownAPIs.insert("Api.AddComment");
    knownAPIs.insert("Api.GetComments");
    knownAPIs.insert("Api.GetAllComments");
    knownAPIs.insert("Api.GetCommentById");
    
    // Named Ranges
    knownAPIs.insert("Api.AddDefName");
    knownAPIs.insert("Api.GetDefName");
    
    // Custom Functions
    knownAPIs.insert("Api.AddCustomFunction");
    knownAPIs.insert("Api.AddCustomFunctionLibrary");
    knownAPIs.insert("Api.ClearCustomFunctions");
    knownAPIs.insert("Api.RemoveCustomFunction");
    
    // Pivot Tables
    knownAPIs.insert("Api.GetAllPivotTables");
    knownAPIs.insert("Api.GetPivotByName");
    knownAPIs.insert("Api.InsertPivotNewWorksheet");
    knownAPIs.insert("Api.InsertPivotExistingWorksheet");
    knownAPIs.insert("Api.RefreshAllPivots");
    
    // Ranges & Selection
    knownAPIs.insert("Api.GetRange");
    knownAPIs.insert("Api.GetSelection");
    knownAPIs.insert("Api.Intersect");
    
    // Formatting & Colors
    knownAPIs.insert("Api.CreateBlipFill");
    knownAPIs.insert("Api.CreateColorByName");
    knownAPIs.insert("Api.CreateColorFromRGB");
    knownAPIs.insert("Api.CreateGradientStop");
    knownAPIs.insert("Api.CreateLinearGradientFill");
    knownAPIs.insert("Api.CreateNoFill");
    knownAPIs.insert("Api.CreatePatternFill");
    knownAPIs.insert("Api.CreatePresetColor");
    knownAPIs.insert("Api.CreateRGBColor");
    knownAPIs.insert("Api.CreateRadialGradientFill");
    knownAPIs.insert("Api.CreateSchemeColor");
    knownAPIs.insert("Api.CreateSolidFill");
    knownAPIs.insert("Api.CreateStroke");
    
    // Text & Typography
    knownAPIs.insert("Api.CreateBullet");
    knownAPIs.insert("Api.CreateNumbering");
    knownAPIs.insert("Api.CreateParagraph");
    knownAPIs.insert("Api.CreateRun");
    knownAPIs.insert("Api.CreateSlide");
    knownAPIs.insert("Api.CreateTextPr");
    
    // Utility & System
    knownAPIs.insert("Api.CreateNewHistoryPoint");
    knownAPIs.insert("Api.Format");
    knownAPIs.insert("Api.GetCore");
    knownAPIs.insert("Api.GetCustomProperties");
    knownAPIs.insert("Api.GetDocumentInfo");
    knownAPIs.insert("Api.GetFreezePanesType");
    knownAPIs.insert("Api.GetFullName");
    knownAPIs.insert("Api.GetLocale");
    knownAPIs.insert("Api.GetMailMergeData");
    knownAPIs.insert("Api.GetReferenceStyle");
    knownAPIs.insert("Api.GetThemesColors");
    knownAPIs.insert("Api.GetWorksheetFunction");
    knownAPIs.insert("Api.RecalculateAllFormulas");
    knownAPIs.insert("Api.ReplaceTextSmart");
    knownAPIs.insert("Api.SetFreezePanesType");
    knownAPIs.insert("Api.SetLocale");
    knownAPIs.insert("Api.SetReferenceStyle");
    knownAPIs.insert("Api.SetThemeColors");
    
    // Events
    knownAPIs.insert("Api.OnDocumentReady");
    knownAPIs.insert("Api.attachEvent");
    knownAPIs.insert("Api.detachEvent");
    knownAPIs.insert("Api.onWorksheetChange");
    
    // ApiRange methods - Value & Data Operations
    knownAPIs.insert("ApiRange.SetValue");
    knownAPIs.insert("ApiRange.GetValue");
    knownAPIs.insert("ApiRange.GetValue2");
    knownAPIs.insert("ApiRange.GetText");
    knownAPIs.insert("ApiRange.Clear");
    knownAPIs.insert("ApiRange.Copy");
    knownAPIs.insert("ApiRange.Cut");
    knownAPIs.insert("ApiRange.Paste");
    knownAPIs.insert("ApiRange.PasteSpecial");
    knownAPIs.insert("ApiRange.Delete");
    knownAPIs.insert("ApiRange.Insert");
    
    // ApiRange methods - Formatting
    knownAPIs.insert("ApiRange.SetFontColor");
    knownAPIs.insert("ApiRange.SetFontName");
    knownAPIs.insert("ApiRange.SetFontSize");
    knownAPIs.insert("ApiRange.SetFillColor");
    knownAPIs.insert("ApiRange.SetBold");
    knownAPIs.insert("ApiRange.SetItalic");
    knownAPIs.insert("ApiRange.SetUnderline");
    knownAPIs.insert("ApiRange.SetStrikeout");
    knownAPIs.insert("ApiRange.SetBorders");
    knownAPIs.insert("ApiRange.SetNumberFormat");
    knownAPIs.insert("ApiRange.GetNumberFormat");
    knownAPIs.insert("ApiRange.GetFillColor");
    
    // ApiRange methods - Alignment & Layout
    knownAPIs.insert("ApiRange.SetAlignHorizontal");
    knownAPIs.insert("ApiRange.SetAlignVertical");
    knownAPIs.insert("ApiRange.SetOrientation");
    knownAPIs.insert("ApiRange.GetOrientation");
    knownAPIs.insert("ApiRange.SetWrap");
    knownAPIs.insert("ApiRange.GetWrapText");
    knownAPIs.insert("ApiRange.SetRowHeight");
    knownAPIs.insert("ApiRange.GetRowHeight");
    knownAPIs.insert("ApiRange.SetColumnWidth");
    knownAPIs.insert("ApiRange.GetColumnWidth");
    knownAPIs.insert("ApiRange.AutoFit");
    knownAPIs.insert("ApiRange.SetHidden");
    knownAPIs.insert("ApiRange.GetHidden");
    
    // ApiRange methods - Formulas & Functions
    knownAPIs.insert("ApiRange.SetFormulaArray");
    knownAPIs.insert("ApiRange.GetFormula");
    knownAPIs.insert("ApiRange.GetFormulaArray");
    
    // ApiRange methods - Navigation & Selection
    knownAPIs.insert("ApiRange.Select");
    knownAPIs.insert("ApiRange.End");
    knownAPIs.insert("ApiRange.GetAddress");
    knownAPIs.insert("ApiRange.SetOffset");
    
    // ApiRange methods - Structure & Relationships
    knownAPIs.insert("ApiRange.Merge");
    knownAPIs.insert("ApiRange.UnMerge");
    knownAPIs.insert("ApiRange.GetCells");
    knownAPIs.insert("ApiRange.GetRows");
    knownAPIs.insert("ApiRange.GetCols");
    knownAPIs.insert("ApiRange.GetRow");
    knownAPIs.insert("ApiRange.GetCol");
    knownAPIs.insert("ApiRange.GetCount");
    knownAPIs.insert("ApiRange.GetWorksheet");
    knownAPIs.insert("ApiRange.GetAreas");
    
    // ApiRange methods - Search & Replace
    knownAPIs.insert("ApiRange.Find");
    knownAPIs.insert("ApiRange.FindNext");
    knownAPIs.insert("ApiRange.FindPrevious");
    knownAPIs.insert("ApiRange.Replace");
    knownAPIs.insert("ApiRange.ForEach");
    
    // ApiRange methods - Comments & Annotations
    knownAPIs.insert("ApiRange.AddComment");
    knownAPIs.insert("ApiRange.GetComment");
    
    // ApiRange methods - Advanced Features
    knownAPIs.insert("ApiRange.SetAutoFilter");
    knownAPIs.insert("ApiRange.SetSort");
    knownAPIs.insert("ApiRange.GetCharacters");
    knownAPIs.insert("ApiRange.GetDefName");
    knownAPIs.insert("ApiRange.GetPivotTable");
    knownAPIs.insert("ApiRange.GetClassType");
    
    // ApiWorksheet methods - Drawing & Graphics
    knownAPIs.insert("ApiWorksheet.AddShape");
    knownAPIs.insert("ApiWorksheet.AddChart");
    knownAPIs.insert("ApiWorksheet.AddImage");
    knownAPIs.insert("ApiWorksheet.AddWordArt");
    knownAPIs.insert("ApiWorksheet.AddOleObject");
    knownAPIs.insert("ApiWorksheet.GetAllShapes");
    knownAPIs.insert("ApiWorksheet.GetAllCharts");
    knownAPIs.insert("ApiWorksheet.GetAllImages");
    knownAPIs.insert("ApiWorksheet.GetAllDrawings");
    knownAPIs.insert("ApiWorksheet.GetAllOleObjects");
    knownAPIs.insert("ApiWorksheet.ReplaceCurrentImage");
    
    // ApiWorksheet methods - Range & Cell Operations
    knownAPIs.insert("ApiWorksheet.GetRange");
    knownAPIs.insert("ApiWorksheet.GetRangeByNumber");
    knownAPIs.insert("ApiWorksheet.GetActiveCell");
    knownAPIs.insert("ApiWorksheet.GetSelection");
    knownAPIs.insert("ApiWorksheet.GetUsedRange");
    knownAPIs.insert("ApiWorksheet.GetCells");
    knownAPIs.insert("ApiWorksheet.GetRows");
    knownAPIs.insert("ApiWorksheet.GetCols");
    knownAPIs.insert("ApiWorksheet.Paste");
    
    // ApiWorksheet methods - Worksheet Properties
    knownAPIs.insert("ApiWorksheet.GetName");
    knownAPIs.insert("ApiWorksheet.SetName");
    knownAPIs.insert("ApiWorksheet.GetIndex");
    knownAPIs.insert("ApiWorksheet.GetVisible");
    knownAPIs.insert("ApiWorksheet.SetVisible");
    knownAPIs.insert("ApiWorksheet.SetActive");
    knownAPIs.insert("ApiWorksheet.Delete");
    knownAPIs.insert("ApiWorksheet.Move");
    knownAPIs.insert("ApiWorksheet.GetClassType");
    
    // ApiWorksheet methods - Page & Print Setup
    knownAPIs.insert("ApiWorksheet.GetPageOrientation");
    knownAPIs.insert("ApiWorksheet.SetPageOrientation");
    knownAPIs.insert("ApiWorksheet.GetPrintGridlines");
    knownAPIs.insert("ApiWorksheet.SetPrintGridlines");
    knownAPIs.insert("ApiWorksheet.GetPrintHeadings");
    knownAPIs.insert("ApiWorksheet.SetPrintHeadings");
    knownAPIs.insert("ApiWorksheet.GetTopMargin");
    knownAPIs.insert("ApiWorksheet.SetTopMargin");
    knownAPIs.insert("ApiWorksheet.GetBottomMargin");
    knownAPIs.insert("ApiWorksheet.SetBottomMargin");
    knownAPIs.insert("ApiWorksheet.GetLeftMargin");
    knownAPIs.insert("ApiWorksheet.SetLeftMargin");
    knownAPIs.insert("ApiWorksheet.GetRightMargin");
    knownAPIs.insert("ApiWorksheet.SetRightMargin");
    
    // ApiWorksheet methods - Display & Layout
    knownAPIs.insert("ApiWorksheet.SetDisplayGridlines");
    knownAPIs.insert("ApiWorksheet.SetDisplayHeadings");
    knownAPIs.insert("ApiWorksheet.SetRowHeight");
    knownAPIs.insert("ApiWorksheet.SetColumnWidth");
    knownAPIs.insert("ApiWorksheet.GetFreezePanes");
    
    // ApiWorksheet methods - Named Ranges & Definitions
    knownAPIs.insert("ApiWorksheet.AddDefName");
    knownAPIs.insert("ApiWorksheet.GetDefName");
    knownAPIs.insert("ApiWorksheet.GetDefNames");
    
    // ApiWorksheet methods - Comments
    knownAPIs.insert("ApiWorksheet.GetComments");
    
    // ApiWorksheet methods - Pivot Tables
    knownAPIs.insert("ApiWorksheet.GetAllPivotTables");
    knownAPIs.insert("ApiWorksheet.GetPivotByName");
    knownAPIs.insert("ApiWorksheet.RefreshAllPivots");
    
    // ApiWorksheet methods - Protection
    knownAPIs.insert("ApiWorksheet.AddProtectedRange");
    knownAPIs.insert("ApiWorksheet.GetAllProtectedRanges");
    knownAPIs.insert("ApiWorksheet.GetProtectedRange");
    
    // ApiWorksheet methods - Formatting & Tables
    knownAPIs.insert("ApiWorksheet.FormatAsTable");
    knownAPIs.insert("ApiWorksheet.SetHyperlink");
    
    // ApiWorksheetFunction methods - Additional functions from analysis
    knownAPIs.insert("ApiWorksheetFunction.COLUMNS");
    knownAPIs.insert("ApiWorksheetFunction.COLUMN");
    
    // ApiChart methods - Series Management
    knownAPIs.insert("ApiChart.AddSeria");
    knownAPIs.insert("ApiChart.RemoveSeria");
    knownAPIs.insert("ApiChart.GetSeries");
    knownAPIs.insert("ApiChart.GetAllSeries");
    knownAPIs.insert("ApiChart.SetSeriaName");
    knownAPIs.insert("ApiChart.SetSeriaValues");
    knownAPIs.insert("ApiChart.SetSeriaXValues");
    
    // ApiChart methods - Chart Appearance
    knownAPIs.insert("ApiChart.SetTitle");
    knownAPIs.insert("ApiChart.SetTitleFill");
    knownAPIs.insert("ApiChart.SetTitleOutLine");
    knownAPIs.insert("ApiChart.ApplyChartStyle");
    knownAPIs.insert("ApiChart.SetSeriesFill");
    knownAPIs.insert("ApiChart.SetSeriesOutLine");
    
    // ApiChart methods - Data Points & Markers
    knownAPIs.insert("ApiChart.SetDataPointFill");
    knownAPIs.insert("ApiChart.SetDataPointOutLine");
    knownAPIs.insert("ApiChart.SetMarkerFill");
    knownAPIs.insert("ApiChart.SetMarkerOutLine");
    
    // ApiChart methods - Axes Configuration
    knownAPIs.insert("ApiChart.SetHorAxisTitle");
    knownAPIs.insert("ApiChart.SetVerAxisTitle");
    knownAPIs.insert("ApiChart.SetHorAxisOrientation");
    knownAPIs.insert("ApiChart.SetVerAxisOrientation");
    knownAPIs.insert("ApiChart.SetHorAxisLablesFontSize");
    knownAPIs.insert("ApiChart.SetVertAxisLablesFontSize");
    knownAPIs.insert("ApiChart.SetHorAxisMajorTickMark");
    knownAPIs.insert("ApiChart.SetHorAxisMinorTickMark");
    knownAPIs.insert("ApiChart.SetVertAxisMajorTickMark");
    knownAPIs.insert("ApiChart.SetVertAxisMinorTickMark");
    knownAPIs.insert("ApiChart.SetHorAxisTickLabelPosition");
    knownAPIs.insert("ApiChart.SetVertAxisTickLabelPosition");
    knownAPIs.insert("ApiChart.SetAxieNumFormat");
    
    // ApiChart methods - Gridlines
    knownAPIs.insert("ApiChart.SetMajorHorizontalGridlines");
    knownAPIs.insert("ApiChart.SetMajorVerticalGridlines");
    knownAPIs.insert("ApiChart.SetMinorHorizontalGridlines");
    knownAPIs.insert("ApiChart.SetMinorVerticalGridlines");
    
    // ApiChart methods - Legend
    knownAPIs.insert("ApiChart.SetLegendPos");
    knownAPIs.insert("ApiChart.SetLegendFill");
    knownAPIs.insert("ApiChart.SetLegendOutLine");
    knownAPIs.insert("ApiChart.SetLegendFontSize");
    
    // ApiChart methods - Plot Area
    knownAPIs.insert("ApiChart.SetPlotAreaFill");
    knownAPIs.insert("ApiChart.SetPlotAreaOutLine");
    
    // ApiChart methods - Data Labels
    knownAPIs.insert("ApiChart.SetShowDataLabels");
    knownAPIs.insert("ApiChart.SetShowPointDataLabel");
    
    // ApiChart methods - Data Configuration
    knownAPIs.insert("ApiChart.SetCatFormula");
    knownAPIs.insert("ApiChart.GetClassType");
    
    // ApiComment methods
    knownAPIs.insert("ApiComment.AddReply");
    knownAPIs.insert("ApiComment.GetReply");
    knownAPIs.insert("ApiComment.GetRepliesCount");
    knownAPIs.insert("ApiComment.RemoveReplies");
    knownAPIs.insert("ApiComment.GetText");
    knownAPIs.insert("ApiComment.SetText");
    knownAPIs.insert("ApiComment.GetAuthorName");
    knownAPIs.insert("ApiComment.SetAuthorName");
    knownAPIs.insert("ApiComment.GetId");
    knownAPIs.insert("ApiComment.GetUserId");
    knownAPIs.insert("ApiComment.SetUserId");
    knownAPIs.insert("ApiComment.GetQuoteText");
    knownAPIs.insert("ApiComment.GetTime");
    knownAPIs.insert("ApiComment.SetTime");
    knownAPIs.insert("ApiComment.GetTimeUTC");
    knownAPIs.insert("ApiComment.SetTimeUTC");
    knownAPIs.insert("ApiComment.IsSolved");
    knownAPIs.insert("ApiComment.SetSolved");
    knownAPIs.insert("ApiComment.Delete");
    knownAPIs.insert("ApiComment.GetClassType");
    
    // ApiCommentReply methods
    knownAPIs.insert("ApiCommentReply.GetAuthorName");
    knownAPIs.insert("ApiCommentReply.SetAuthorName");
    knownAPIs.insert("ApiCommentReply.GetClassType");
    knownAPIs.insert("ApiCommentReply.GetText");
    knownAPIs.insert("ApiCommentReply.SetText");
    knownAPIs.insert("ApiCommentReply.GetTime");
    knownAPIs.insert("ApiCommentReply.SetTime");
    knownAPIs.insert("ApiCommentReply.GetTimeUTC");
    knownAPIs.insert("ApiCommentReply.SetTimeUTC");
    knownAPIs.insert("ApiCommentReply.GetUserId");
    knownAPIs.insert("ApiCommentReply.SetUserId");
    
    // ApiFont methods
    knownAPIs.insert("ApiFont.GetBold");
    knownAPIs.insert("ApiFont.SetBold");
    knownAPIs.insert("ApiFont.GetItalic");
    knownAPIs.insert("ApiFont.SetItalic");
    knownAPIs.insert("ApiFont.GetUnderline");
    knownAPIs.insert("ApiFont.SetUnderline");
    knownAPIs.insert("ApiFont.GetStrikethrough");
    knownAPIs.insert("ApiFont.SetStrikethrough");
    knownAPIs.insert("ApiFont.GetSubscript");
    knownAPIs.insert("ApiFont.SetSubscript");
    knownAPIs.insert("ApiFont.GetSuperscript");
    knownAPIs.insert("ApiFont.SetSuperscript");
    knownAPIs.insert("ApiFont.GetName");
    knownAPIs.insert("ApiFont.SetName");
    knownAPIs.insert("ApiFont.GetSize");
    knownAPIs.insert("ApiFont.SetSize");
    knownAPIs.insert("ApiFont.GetColor");
    knownAPIs.insert("ApiFont.SetColor");
    knownAPIs.insert("ApiFont.GetParent");
    
    // ApiPivotTable methods (comprehensive list from analysis)
    knownAPIs.insert("ApiPivotTable.AddDataField");
    knownAPIs.insert("ApiPivotTable.AddFields");
    knownAPIs.insert("ApiPivotTable.ClearAllFilters");
    knownAPIs.insert("ApiPivotTable.ClearTable");
    knownAPIs.insert("ApiPivotTable.GetColumnFields");
    knownAPIs.insert("ApiPivotTable.GetColumnGrand");
    knownAPIs.insert("ApiPivotTable.SetColumnGrand");
    knownAPIs.insert("ApiPivotTable.GetColumnRange");
    knownAPIs.insert("ApiPivotTable.GetData");
    knownAPIs.insert("ApiPivotTable.GetDataBodyRange");
    knownAPIs.insert("ApiPivotTable.GetDataFields");
    knownAPIs.insert("ApiPivotTable.GetDescription");
    knownAPIs.insert("ApiPivotTable.SetDescription");
    knownAPIs.insert("ApiPivotTable.GetDisplayFieldCaptions");
    knownAPIs.insert("ApiPivotTable.SetDisplayFieldCaptions");
    knownAPIs.insert("ApiPivotTable.GetDisplayFieldsInReportFilterArea");
    knownAPIs.insert("ApiPivotTable.SetDisplayFieldsInReportFilterArea");
    knownAPIs.insert("ApiPivotTable.GetGrandTotalName");
    knownAPIs.insert("ApiPivotTable.SetGrandTotalName");
    knownAPIs.insert("ApiPivotTable.GetHiddenFields");
    knownAPIs.insert("ApiPivotTable.GetName");
    knownAPIs.insert("ApiPivotTable.SetName");
    knownAPIs.insert("ApiPivotTable.GetPageFields");
    knownAPIs.insert("ApiPivotTable.GetParent");
    knownAPIs.insert("ApiPivotTable.GetPivotData");
    knownAPIs.insert("ApiPivotTable.GetPivotFields");
    knownAPIs.insert("ApiPivotTable.GetRowFields");
    knownAPIs.insert("ApiPivotTable.GetRowGrand");
    knownAPIs.insert("ApiPivotTable.SetRowGrand");
    knownAPIs.insert("ApiPivotTable.GetRowRange");
    knownAPIs.insert("ApiPivotTable.GetSource");
    knownAPIs.insert("ApiPivotTable.SetSource");
    knownAPIs.insert("ApiPivotTable.GetStyleName");
    knownAPIs.insert("ApiPivotTable.SetStyleName");
    knownAPIs.insert("ApiPivotTable.GetTableRange1");
    knownAPIs.insert("ApiPivotTable.GetTableRange2");
    knownAPIs.insert("ApiPivotTable.GetTableStyleColumnHeaders");
    knownAPIs.insert("ApiPivotTable.SetTableStyleColumnHeaders");
    knownAPIs.insert("ApiPivotTable.GetTableStyleColumnStripes");
    knownAPIs.insert("ApiPivotTable.SetTableStyleColumnStripes");
    knownAPIs.insert("ApiPivotTable.GetTableStyleRowHeaders");
    knownAPIs.insert("ApiPivotTable.SetTableStyleRowHeaders");
    knownAPIs.insert("ApiPivotTable.GetTableStyleRowStripes");
    knownAPIs.insert("ApiPivotTable.SetTableStyleRowStripes");
    knownAPIs.insert("ApiPivotTable.GetTitle");
    knownAPIs.insert("ApiPivotTable.SetTitle");
    knownAPIs.insert("ApiPivotTable.GetVisibleFields");
    knownAPIs.insert("ApiPivotTable.MoveField");
    knownAPIs.insert("ApiPivotTable.PivotValueCell");
    knownAPIs.insert("ApiPivotTable.RefreshTable");
    knownAPIs.insert("ApiPivotTable.RemoveField");
    knownAPIs.insert("ApiPivotTable.Select");
    knownAPIs.insert("ApiPivotTable.SetLayoutBlankLine");
    knownAPIs.insert("ApiPivotTable.SetLayoutSubtotals");
    knownAPIs.insert("ApiPivotTable.SetRepeatAllLabels");
    knownAPIs.insert("ApiPivotTable.SetRowAxisLayout");
    knownAPIs.insert("ApiPivotTable.SetSubtotalLocation");
    knownAPIs.insert("ApiPivotTable.ShowDetails");
    knownAPIs.insert("ApiPivotTable.Update");
    
    // ApiPivotField methods
    knownAPIs.insert("ApiPivotField.ClearAllFilters");
    knownAPIs.insert("ApiPivotField.ClearLabelFilters");
    knownAPIs.insert("ApiPivotField.ClearManualFilters");
    knownAPIs.insert("ApiPivotField.ClearValueFilters");
    knownAPIs.insert("ApiPivotField.GetCaption");
    knownAPIs.insert("ApiPivotField.SetCaption");
    knownAPIs.insert("ApiPivotField.GetCurrentPage");
    knownAPIs.insert("ApiPivotField.GetDragToColumn");
    knownAPIs.insert("ApiPivotField.SetDragToColumn");
    knownAPIs.insert("ApiPivotField.GetDragToData");
    knownAPIs.insert("ApiPivotField.SetDragToData");
    knownAPIs.insert("ApiPivotField.GetDragToPage");
    knownAPIs.insert("ApiPivotField.SetDragToPage");
    knownAPIs.insert("ApiPivotField.GetDragToRow");
    knownAPIs.insert("ApiPivotField.SetDragToRow");
    knownAPIs.insert("ApiPivotField.GetIndex");
    knownAPIs.insert("ApiPivotField.GetLayoutBlankLine");
    knownAPIs.insert("ApiPivotField.SetLayoutBlankLine");
    knownAPIs.insert("ApiPivotField.GetLayoutCompactRow");
    knownAPIs.insert("ApiPivotField.SetLayoutCompactRow");
    knownAPIs.insert("ApiPivotField.GetLayoutForm");
    knownAPIs.insert("ApiPivotField.SetLayoutForm");
    knownAPIs.insert("ApiPivotField.GetLayoutPageBreak");
    knownAPIs.insert("ApiPivotField.SetLayoutPageBreak");
    knownAPIs.insert("ApiPivotField.GetLayoutSubtotalLocation");
    knownAPIs.insert("ApiPivotField.SetLayoutSubtotalLocation");
    knownAPIs.insert("ApiPivotField.GetLayoutSubtotals");
    knownAPIs.insert("ApiPivotField.SetLayoutSubtotals");
    knownAPIs.insert("ApiPivotField.GetName");
    knownAPIs.insert("ApiPivotField.SetName");
    knownAPIs.insert("ApiPivotField.GetOrientation");
    knownAPIs.insert("ApiPivotField.SetOrientation");
    knownAPIs.insert("ApiPivotField.GetParent");
    knownAPIs.insert("ApiPivotField.GetPivotItems");
    knownAPIs.insert("ApiPivotField.GetPosition");
    knownAPIs.insert("ApiPivotField.SetPosition");
    knownAPIs.insert("ApiPivotField.GetRepeatLabels");
    knownAPIs.insert("ApiPivotField.SetRepeatLabels");
    knownAPIs.insert("ApiPivotField.GetShowAllItems");
    knownAPIs.insert("ApiPivotField.SetShowAllItems");
    knownAPIs.insert("ApiPivotField.GetShowingInAxis");
    knownAPIs.insert("ApiPivotField.GetSourceName");
    knownAPIs.insert("ApiPivotField.GetSubtotalName");
    knownAPIs.insert("ApiPivotField.SetSubtotalName");
    knownAPIs.insert("ApiPivotField.GetSubtotals");
    knownAPIs.insert("ApiPivotField.SetSubtotals");
    knownAPIs.insert("ApiPivotField.GetTable");
    knownAPIs.insert("ApiPivotField.GetValue");
    knownAPIs.insert("ApiPivotField.SetValue");
    knownAPIs.insert("ApiPivotField.Move");
    knownAPIs.insert("ApiPivotField.Remove");
    
    // ApiPivotDataField methods
    knownAPIs.insert("ApiPivotDataField.GetCaption");
    knownAPIs.insert("ApiPivotDataField.SetCaption");
    knownAPIs.insert("ApiPivotDataField.GetFunction");
    knownAPIs.insert("ApiPivotDataField.SetFunction");
    knownAPIs.insert("ApiPivotDataField.GetIndex");
    knownAPIs.insert("ApiPivotDataField.GetName");
    knownAPIs.insert("ApiPivotDataField.SetName");
    knownAPIs.insert("ApiPivotDataField.GetNumberFormat");
    knownAPIs.insert("ApiPivotDataField.SetNumberFormat");
    knownAPIs.insert("ApiPivotDataField.GetOrientation");
    knownAPIs.insert("ApiPivotDataField.GetPivotField");
    knownAPIs.insert("ApiPivotDataField.GetPosition");
    knownAPIs.insert("ApiPivotDataField.SetPosition");
    knownAPIs.insert("ApiPivotDataField.GetValue");
    knownAPIs.insert("ApiPivotDataField.SetValue");
    knownAPIs.insert("ApiPivotDataField.Move");
    knownAPIs.insert("ApiPivotDataField.Remove");
    
    // ApiPivotItem methods
    knownAPIs.insert("ApiPivotItem.GetCaption");
    knownAPIs.insert("ApiPivotItem.GetName");
    knownAPIs.insert("ApiPivotItem.GetParent");
    knownAPIs.insert("ApiPivotItem.GetValue");
    
    // ApiProtectedRange methods
    knownAPIs.insert("ApiProtectedRange.AddUser");
    knownAPIs.insert("ApiProtectedRange.DeleteUser");
    knownAPIs.insert("ApiProtectedRange.GetAllUsers");
    knownAPIs.insert("ApiProtectedRange.GetUser");
    knownAPIs.insert("ApiProtectedRange.SetAnyoneType");
    knownAPIs.insert("ApiProtectedRange.SetRange");
    knownAPIs.insert("ApiProtectedRange.SetTitle");
    
    // ApiProtectedRangeUserInfo methods
    knownAPIs.insert("ApiProtectedRangeUserInfo.GetId");
    knownAPIs.insert("ApiProtectedRangeUserInfo.GetName");
    knownAPIs.insert("ApiProtectedRangeUserInfo.GetType");
    
    // Additional common API objects
    knownAPIs.insert("ApiDrawing.GetHeight");
    knownAPIs.insert("ApiDrawing.GetWidth");
    knownAPIs.insert("ApiDrawing.SetSize");
    knownAPIs.insert("ApiDrawing.SetPosition");
    knownAPIs.insert("ApiDrawing.GetClassType");
    knownAPIs.insert("ApiDrawing.GetLockValue");
    knownAPIs.insert("ApiDrawing.SetLockValue");
    knownAPIs.insert("ApiDrawing.GetRotation");
    knownAPIs.insert("ApiDrawing.SetRotation");
    knownAPIs.insert("ApiDrawing.GetParentSheet");
    
    knownAPIs.insert("ApiShape.GetContent");
    knownAPIs.insert("ApiShape.GetDocContent");
    knownAPIs.insert("ApiShape.SetVerticalTextAlign");
    knownAPIs.insert("ApiShape.GetClassType");
    
    knownAPIs.insert("ApiName.GetName");
    knownAPIs.insert("ApiName.SetName");
    knownAPIs.insert("ApiName.GetRefersTo");
    knownAPIs.insert("ApiName.SetRefersTo");
    knownAPIs.insert("ApiName.GetRefersToRange");
    knownAPIs.insert("ApiName.Delete");
    
    // ApiFreezePanes methods
    knownAPIs.insert("ApiFreezePanes.FreezeAt");
    knownAPIs.insert("ApiFreezePanes.FreezeColumns");
    knownAPIs.insert("ApiFreezePanes.FreezeRows");
    knownAPIs.insert("ApiFreezePanes.GetLocation");
    knownAPIs.insert("ApiFreezePanes.Unfreeze");
    
    // ApiOleObject methods
    knownAPIs.insert("ApiOleObject.GetApplicationId");
    knownAPIs.insert("ApiOleObject.SetApplicationId");
    knownAPIs.insert("ApiOleObject.GetData");
    knownAPIs.insert("ApiOleObject.SetData");
    knownAPIs.insert("ApiOleObject.GetClassType");
    
    // ApiImage methods
    knownAPIs.insert("ApiImage.GetClassType");
    
    // ApiCore methods (document properties)
    knownAPIs.insert("ApiCore.GetCategory");
    knownAPIs.insert("ApiCore.SetCategory");
    knownAPIs.insert("ApiCore.GetClassType");
    knownAPIs.insert("ApiCore.GetContentStatus");
    knownAPIs.insert("ApiCore.SetContentStatus");
    knownAPIs.insert("ApiCore.GetCreated");
    knownAPIs.insert("ApiCore.SetCreated");
    knownAPIs.insert("ApiCore.GetCreator");
    knownAPIs.insert("ApiCore.SetCreator");
    knownAPIs.insert("ApiCore.GetDescription");
    knownAPIs.insert("ApiCore.SetDescription");
    knownAPIs.insert("ApiCore.GetIdentifier");
    knownAPIs.insert("ApiCore.SetIdentifier");
    knownAPIs.insert("ApiCore.GetKeywords");
    knownAPIs.insert("ApiCore.SetKeywords");
    knownAPIs.insert("ApiCore.GetLanguage");
    knownAPIs.insert("ApiCore.SetLanguage");
    knownAPIs.insert("ApiCore.GetLastModifiedBy");
    knownAPIs.insert("ApiCore.SetLastModifiedBy");
    knownAPIs.insert("ApiCore.GetLastPrinted");
    knownAPIs.insert("ApiCore.SetLastPrinted");
    knownAPIs.insert("ApiCore.GetModified");
    knownAPIs.insert("ApiCore.SetModified");
    knownAPIs.insert("ApiCore.GetRevision");
    knownAPIs.insert("ApiCore.SetRevision");
    knownAPIs.insert("ApiCore.GetSubject");
    knownAPIs.insert("ApiCore.SetSubject");
    knownAPIs.insert("ApiCore.GetTitle");
    knownAPIs.insert("ApiCore.SetTitle");
    knownAPIs.insert("ApiCore.GetVersion");
    knownAPIs.insert("ApiCore.SetVersion");
    
    // ApiCustomProperties methods
    knownAPIs.insert("ApiCustomProperties.Add");
    knownAPIs.insert("ApiCustomProperties.Get");
    knownAPIs.insert("ApiCustomProperties.GetClassType");
    
    // ApiDocumentContent methods
    knownAPIs.insert("ApiDocumentContent.AddElement");
    knownAPIs.insert("ApiDocumentContent.GetClassType");
    knownAPIs.insert("ApiDocumentContent.GetElement");
    knownAPIs.insert("ApiDocumentContent.GetElementsCount");
    knownAPIs.insert("ApiDocumentContent.Push");
    knownAPIs.insert("ApiDocumentContent.RemoveAllElements");
    knownAPIs.insert("ApiDocumentContent.RemoveElement");
    
    // ApiParaPr methods (paragraph properties)
    knownAPIs.insert("ApiParaPr.GetClassType");
    knownAPIs.insert("ApiParaPr.GetIndFirstLine");
    knownAPIs.insert("ApiParaPr.SetIndFirstLine");
    knownAPIs.insert("ApiParaPr.GetIndLeft");
    knownAPIs.insert("ApiParaPr.SetIndLeft");
    knownAPIs.insert("ApiParaPr.GetIndRight");
    knownAPIs.insert("ApiParaPr.SetIndRight");
    knownAPIs.insert("ApiParaPr.GetJc");
    knownAPIs.insert("ApiParaPr.SetJc");
    knownAPIs.insert("ApiParaPr.GetSpacingAfter");
    knownAPIs.insert("ApiParaPr.SetSpacingAfter");
    knownAPIs.insert("ApiParaPr.GetSpacingBefore");
    knownAPIs.insert("ApiParaPr.SetSpacingBefore");
    knownAPIs.insert("ApiParaPr.GetSpacingLineRule");
    knownAPIs.insert("ApiParaPr.GetSpacingLineValue");
    knownAPIs.insert("ApiParaPr.SetBullet");
    knownAPIs.insert("ApiParaPr.SetSpacingLine");
    knownAPIs.insert("ApiParaPr.SetTabs");
    
    // ApiParagraph methods
    knownAPIs.insert("ApiParagraph.AddElement");
    knownAPIs.insert("ApiParagraph.AddLineBreak");
    knownAPIs.insert("ApiParagraph.AddTabStop");
    knownAPIs.insert("ApiParagraph.AddText");
    knownAPIs.insert("ApiParagraph.Copy");
    knownAPIs.insert("ApiParagraph.Delete");
    knownAPIs.insert("ApiParagraph.GetClassType");
    knownAPIs.insert("ApiParagraph.GetElement");
    knownAPIs.insert("ApiParagraph.GetElementsCount");
    knownAPIs.insert("ApiParagraph.GetIndFirstLine");
    knownAPIs.insert("ApiParagraph.SetIndFirstLine");
    knownAPIs.insert("ApiParagraph.GetIndLeft");
    knownAPIs.insert("ApiParagraph.SetIndLeft");
    knownAPIs.insert("ApiParagraph.GetIndRight");
    knownAPIs.insert("ApiParagraph.SetIndRight");
    knownAPIs.insert("ApiParagraph.GetJc");
    knownAPIs.insert("ApiParagraph.SetJc");
    knownAPIs.insert("ApiParagraph.GetNext");
    knownAPIs.insert("ApiParagraph.GetParaPr");
    knownAPIs.insert("ApiParagraph.GetPrevious");
    knownAPIs.insert("ApiParagraph.GetSpacingAfter");
    knownAPIs.insert("ApiParagraph.SetSpacingAfter");
    knownAPIs.insert("ApiParagraph.GetSpacingBefore");
    knownAPIs.insert("ApiParagraph.SetSpacingBefore");
    knownAPIs.insert("ApiParagraph.GetSpacingLineRule");
    knownAPIs.insert("ApiParagraph.GetSpacingLineValue");
    knownAPIs.insert("ApiParagraph.RemoveAllElements");
    knownAPIs.insert("ApiParagraph.RemoveElement");
    knownAPIs.insert("ApiParagraph.SetBullet");
    knownAPIs.insert("ApiParagraph.SetSpacingLine");
    knownAPIs.insert("ApiParagraph.SetTabs");
    
    // ApiRun methods
    knownAPIs.insert("ApiRun.AddLineBreak");
    knownAPIs.insert("ApiRun.AddTabStop");
    knownAPIs.insert("ApiRun.AddText");
    knownAPIs.insert("ApiRun.ClearContent");
    knownAPIs.insert("ApiRun.Copy");
    knownAPIs.insert("ApiRun.Delete");
    knownAPIs.insert("ApiRun.GetClassType");
    knownAPIs.insert("ApiRun.GetFontNames");
    knownAPIs.insert("ApiRun.GetTextPr");
    knownAPIs.insert("ApiRun.RemoveAllElements");
    knownAPIs.insert("ApiRun.SetBold");
    knownAPIs.insert("ApiRun.SetCaps");
    knownAPIs.insert("ApiRun.SetColor");
    knownAPIs.insert("ApiRun.SetDoubleStrikeout");
    knownAPIs.insert("ApiRun.SetFill");
    knownAPIs.insert("ApiRun.SetFontFamily");
    knownAPIs.insert("ApiRun.SetFontSize");
    knownAPIs.insert("ApiRun.SetHighlight");
    knownAPIs.insert("ApiRun.SetItalic");
    knownAPIs.insert("ApiRun.SetLanguage");
    knownAPIs.insert("ApiRun.SetOutLine");
    knownAPIs.insert("ApiRun.SetPosition");
    knownAPIs.insert("ApiRun.SetShd");
    knownAPIs.insert("ApiRun.SetSmallCaps");
    knownAPIs.insert("ApiRun.SetSpacing");
    knownAPIs.insert("ApiRun.SetStrikeout");
    
    // ApiTextPr methods
    knownAPIs.insert("ApiTextPr.GetClassType");
    
    // ApiUniColor methods
    knownAPIs.insert("ApiUniColor.GetClassType");
    
    // ApiSchemeColor methods
    knownAPIs.insert("ApiSchemeColor.GetClassType");
    
    // ApiStroke methods
    knownAPIs.insert("ApiStroke.GetClassType");
    
    // Common class type methods for other objects
    knownAPIs.insert("ApiAreas.GetCount");
    knownAPIs.insert("ApiAreas.GetItem");
    knownAPIs.insert("ApiAreas.GetParent");
    knownAPIs.insert("ApiBullet.GetClassType");
    knownAPIs.insert("ApiCharacters.Delete");
    knownAPIs.insert("ApiCharacters.GetCaption");
    knownAPIs.insert("ApiCharacters.SetCaption");
    knownAPIs.insert("ApiCharacters.GetCount");
    knownAPIs.insert("ApiCharacters.GetFont");
    knownAPIs.insert("ApiCharacters.GetParent");
    knownAPIs.insert("ApiCharacters.GetText");
    knownAPIs.insert("ApiCharacters.SetText");
    knownAPIs.insert("ApiCharacters.Insert");
    knownAPIs.insert("ApiChartSeries.GetChartType");
    knownAPIs.insert("ApiChartSeries.ChangeChartType");
    knownAPIs.insert("ApiChartSeries.GetClassType");
    knownAPIs.insert("ApiColor.GetClassType");
    knownAPIs.insert("ApiColor.GetRGB");
    knownAPIs.insert("ApiRGBColor.GetClassType");
    knownAPIs.insert("ApiPresetColor.GetClassType");
    knownAPIs.insert("ApiFill.GetClassType");
    knownAPIs.insert("ApiGradientStop.GetClassType");
    
    // ApiWorksheetFunction methods (comprehensive list from analysis)
    // Mathematical Functions
    knownAPIs.insert("ApiWorksheetFunction.ABS");
    knownAPIs.insert("ApiWorksheetFunction.ACOS");
    knownAPIs.insert("ApiWorksheetFunction.ASIN");
    knownAPIs.insert("ApiWorksheetFunction.ATAN");
    knownAPIs.insert("ApiWorksheetFunction.ATAN2");
    knownAPIs.insert("ApiWorksheetFunction.AVERAGE");
    knownAPIs.insert("ApiWorksheetFunction.CEILING");
    knownAPIs.insert("ApiWorksheetFunction.COS");
    knownAPIs.insert("ApiWorksheetFunction.COUNT");
    knownAPIs.insert("ApiWorksheetFunction.DEGREES");
    knownAPIs.insert("ApiWorksheetFunction.EXP");
    knownAPIs.insert("ApiWorksheetFunction.FLOOR");
    knownAPIs.insert("ApiWorksheetFunction.LOG");
    knownAPIs.insert("ApiWorksheetFunction.LOG10");
    knownAPIs.insert("ApiWorksheetFunction.MAX");
    knownAPIs.insert("ApiWorksheetFunction.MIN");
    knownAPIs.insert("ApiWorksheetFunction.MOD");
    knownAPIs.insert("ApiWorksheetFunction.PI");
    knownAPIs.insert("ApiWorksheetFunction.POWER");
    knownAPIs.insert("ApiWorksheetFunction.PRODUCT");
    knownAPIs.insert("ApiWorksheetFunction.RADIANS");
    knownAPIs.insert("ApiWorksheetFunction.RAND");
    knownAPIs.insert("ApiWorksheetFunction.RANDBETWEEN");
    knownAPIs.insert("ApiWorksheetFunction.ROUND");
    knownAPIs.insert("ApiWorksheetFunction.ROUNDDOWN");
    knownAPIs.insert("ApiWorksheetFunction.ROUNDUP");
    knownAPIs.insert("ApiWorksheetFunction.SIN");
    knownAPIs.insert("ApiWorksheetFunction.SQRT");
    knownAPIs.insert("ApiWorksheetFunction.SUM");
    knownAPIs.insert("ApiWorksheetFunction.SUMPRODUCT");
    knownAPIs.insert("ApiWorksheetFunction.TAN");
    knownAPIs.insert("ApiWorksheetFunction.TRUNC");
    
    // Statistical Functions
    knownAPIs.insert("ApiWorksheetFunction.AVERAGEIF");
    knownAPIs.insert("ApiWorksheetFunction.AVERAGEIFS");
    knownAPIs.insert("ApiWorksheetFunction.COUNTA");
    knownAPIs.insert("ApiWorksheetFunction.COUNTBLANK");
    knownAPIs.insert("ApiWorksheetFunction.COUNTIF");
    knownAPIs.insert("ApiWorksheetFunction.COUNTIFS");
    knownAPIs.insert("ApiWorksheetFunction.LARGE");
    knownAPIs.insert("ApiWorksheetFunction.MEDIAN");
    knownAPIs.insert("ApiWorksheetFunction.MODE");
    knownAPIs.insert("ApiWorksheetFunction.PERCENTILE");
    knownAPIs.insert("ApiWorksheetFunction.QUARTILE");
    knownAPIs.insert("ApiWorksheetFunction.RANK");
    knownAPIs.insert("ApiWorksheetFunction.SMALL");
    knownAPIs.insert("ApiWorksheetFunction.STDEV");
    knownAPIs.insert("ApiWorksheetFunction.STDEVP");
    knownAPIs.insert("ApiWorksheetFunction.SUMIF");
    knownAPIs.insert("ApiWorksheetFunction.SUMIFS");
    knownAPIs.insert("ApiWorksheetFunction.VAR");
    knownAPIs.insert("ApiWorksheetFunction.VARP");
    
    // Text Functions
    knownAPIs.insert("ApiWorksheetFunction.CHAR");
    knownAPIs.insert("ApiWorksheetFunction.CLEAN");
    knownAPIs.insert("ApiWorksheetFunction.CODE");
    knownAPIs.insert("ApiWorksheetFunction.CONCATENATE");
    knownAPIs.insert("ApiWorksheetFunction.EXACT");
    knownAPIs.insert("ApiWorksheetFunction.FIND");
    knownAPIs.insert("ApiWorksheetFunction.FIXED");
    knownAPIs.insert("ApiWorksheetFunction.LEFT");
    knownAPIs.insert("ApiWorksheetFunction.LEN");
    knownAPIs.insert("ApiWorksheetFunction.LOWER");
    knownAPIs.insert("ApiWorksheetFunction.MID");
    knownAPIs.insert("ApiWorksheetFunction.PROPER");
    knownAPIs.insert("ApiWorksheetFunction.REPLACE");
    knownAPIs.insert("ApiWorksheetFunction.REPT");
    knownAPIs.insert("ApiWorksheetFunction.RIGHT");
    knownAPIs.insert("ApiWorksheetFunction.SEARCH");
    knownAPIs.insert("ApiWorksheetFunction.SUBSTITUTE");
    knownAPIs.insert("ApiWorksheetFunction.TEXT");
    knownAPIs.insert("ApiWorksheetFunction.TRIM");
    knownAPIs.insert("ApiWorksheetFunction.UPPER");
    knownAPIs.insert("ApiWorksheetFunction.VALUE");
    
    // Date/Time Functions
    knownAPIs.insert("ApiWorksheetFunction.DATE");
    knownAPIs.insert("ApiWorksheetFunction.DATEVALUE");
    knownAPIs.insert("ApiWorksheetFunction.DAY");
    knownAPIs.insert("ApiWorksheetFunction.DAYS");
    knownAPIs.insert("ApiWorksheetFunction.DAYS360");
    knownAPIs.insert("ApiWorksheetFunction.EDATE");
    knownAPIs.insert("ApiWorksheetFunction.EOMONTH");
    knownAPIs.insert("ApiWorksheetFunction.HOUR");
    knownAPIs.insert("ApiWorksheetFunction.MINUTE");
    knownAPIs.insert("ApiWorksheetFunction.MONTH");
    knownAPIs.insert("ApiWorksheetFunction.NETWORKDAYS");
    knownAPIs.insert("ApiWorksheetFunction.NOW");
    knownAPIs.insert("ApiWorksheetFunction.SECOND");
    knownAPIs.insert("ApiWorksheetFunction.TIME");
    knownAPIs.insert("ApiWorksheetFunction.TIMEVALUE");
    knownAPIs.insert("ApiWorksheetFunction.TODAY");
    knownAPIs.insert("ApiWorksheetFunction.WEEKDAY");
    knownAPIs.insert("ApiWorksheetFunction.WORKDAY");
    knownAPIs.insert("ApiWorksheetFunction.YEAR");
    knownAPIs.insert("ApiWorksheetFunction.YEARFRAC");
    
    // Logical Functions
    knownAPIs.insert("ApiWorksheetFunction.AND");
    knownAPIs.insert("ApiWorksheetFunction.FALSE");
    knownAPIs.insert("ApiWorksheetFunction.IF");
    knownAPIs.insert("ApiWorksheetFunction.IFERROR");
    knownAPIs.insert("ApiWorksheetFunction.NOT");
    knownAPIs.insert("ApiWorksheetFunction.OR");
    knownAPIs.insert("ApiWorksheetFunction.TRUE");
    
    // Lookup Functions
    knownAPIs.insert("ApiWorksheetFunction.CHOOSE");
    knownAPIs.insert("ApiWorksheetFunction.HLOOKUP");
    knownAPIs.insert("ApiWorksheetFunction.INDEX");
    knownAPIs.insert("ApiWorksheetFunction.INDIRECT");
    knownAPIs.insert("ApiWorksheetFunction.LOOKUP");
    knownAPIs.insert("ApiWorksheetFunction.MATCH");
    knownAPIs.insert("ApiWorksheetFunction.OFFSET");
    knownAPIs.insert("ApiWorksheetFunction.ROW");
    knownAPIs.insert("ApiWorksheetFunction.ROWS");
    knownAPIs.insert("ApiWorksheetFunction.TRANSPOSE");
    knownAPIs.insert("ApiWorksheetFunction.VLOOKUP");
    
    // Financial Functions
    knownAPIs.insert("ApiWorksheetFunction.FV");
    knownAPIs.insert("ApiWorksheetFunction.IPMT");
    knownAPIs.insert("ApiWorksheetFunction.IRR");
    knownAPIs.insert("ApiWorksheetFunction.NPER");
    knownAPIs.insert("ApiWorksheetFunction.NPV");
    knownAPIs.insert("ApiWorksheetFunction.PMT");
    knownAPIs.insert("ApiWorksheetFunction.PPMT");
    knownAPIs.insert("ApiWorksheetFunction.PV");
    knownAPIs.insert("ApiWorksheetFunction.RATE");
    
    // Information Functions
    knownAPIs.insert("ApiWorksheetFunction.CELL");
    knownAPIs.insert("ApiWorksheetFunction.ERROR");
    knownAPIs.insert("ApiWorksheetFunction.INFO");
    knownAPIs.insert("ApiWorksheetFunction.ISBLANK");
    knownAPIs.insert("ApiWorksheetFunction.ISERROR");
    knownAPIs.insert("ApiWorksheetFunction.ISLOGICAL");
    knownAPIs.insert("ApiWorksheetFunction.ISNA");
    knownAPIs.insert("ApiWorksheetFunction.ISNONTEXT");
    knownAPIs.insert("ApiWorksheetFunction.ISNUMBER");
    knownAPIs.insert("ApiWorksheetFunction.ISREF");
    knownAPIs.insert("ApiWorksheetFunction.ISTEXT");
    knownAPIs.insert("ApiWorksheetFunction.N");
    knownAPIs.insert("ApiWorksheetFunction.NA");
    knownAPIs.insert("ApiWorksheetFunction.TYPE");
}

void setupKnownGlobals(std::set<std::string>& knownGlobals) {
    // JavaScript Language Keywords (control flow)
    knownGlobals.insert("if");
    knownGlobals.insert("else");
    knownGlobals.insert("for");
    knownGlobals.insert("while");
    knownGlobals.insert("do");
    knownGlobals.insert("switch");
    knownGlobals.insert("case");
    knownGlobals.insert("default");
    knownGlobals.insert("break");
    knownGlobals.insert("continue");
    knownGlobals.insert("return");
    
    // JavaScript Exception Handling
    knownGlobals.insert("try");
    knownGlobals.insert("catch");
    knownGlobals.insert("finally");
    knownGlobals.insert("throw");
    
    // JavaScript Operators and Keywords
    knownGlobals.insert("typeof");
    knownGlobals.insert("instanceof");
    knownGlobals.insert("in");
    knownGlobals.insert("new");
    knownGlobals.insert("delete");
    knownGlobals.insert("this");
    knownGlobals.insert("super");
    knownGlobals.insert("with");
    
    // JavaScript Variable Declaration
    knownGlobals.insert("var");
    knownGlobals.insert("let");
    knownGlobals.insert("const");
    knownGlobals.insert("function");
    
    // JavaScript Primitive Values
    knownGlobals.insert("undefined");
    knownGlobals.insert("null");
    knownGlobals.insert("true");
    knownGlobals.insert("false");
    knownGlobals.insert("Infinity");
    knownGlobals.insert("NaN");
    
    // JavaScript Built-in Global Functions
    knownGlobals.insert("eval");
    knownGlobals.insert("parseInt");
    knownGlobals.insert("parseFloat");
    knownGlobals.insert("isNaN");
    knownGlobals.insert("isFinite");
    knownGlobals.insert("decodeURI");
    knownGlobals.insert("decodeURIComponent");
    knownGlobals.insert("encodeURI");
    knownGlobals.insert("encodeURIComponent");
    knownGlobals.insert("escape");
    knownGlobals.insert("unescape");
    
    // JavaScript Built-in Constructors
    knownGlobals.insert("Object");
    knownGlobals.insert("Array");
    knownGlobals.insert("String");
    knownGlobals.insert("Number");
    knownGlobals.insert("Boolean");
    knownGlobals.insert("Date");
    knownGlobals.insert("RegExp");
    knownGlobals.insert("Error");
    knownGlobals.insert("EvalError");
    knownGlobals.insert("RangeError");
    knownGlobals.insert("ReferenceError");
    knownGlobals.insert("SyntaxError");
    knownGlobals.insert("TypeError");
    knownGlobals.insert("URIError");
    knownGlobals.insert("Function");
    
    // JavaScript Built-in Objects
    knownGlobals.insert("Math");
    knownGlobals.insert("JSON");
    knownGlobals.insert("Reflect");
    knownGlobals.insert("Proxy");
    knownGlobals.insert("Promise");
    knownGlobals.insert("Symbol");
    knownGlobals.insert("Map");
    knownGlobals.insert("Set");
    knownGlobals.insert("WeakMap");
    knownGlobals.insert("WeakSet");
    knownGlobals.insert("ArrayBuffer");
    knownGlobals.insert("DataView");
    knownGlobals.insert("Int8Array");
    knownGlobals.insert("Uint8Array");
    knownGlobals.insert("Int16Array");
    knownGlobals.insert("Uint16Array");
    knownGlobals.insert("Int32Array");
    knownGlobals.insert("Uint32Array");
    knownGlobals.insert("Float32Array");
    knownGlobals.insert("Float64Array");
    
    // Console and Environment
    knownGlobals.insert("console");
    knownGlobals.insert("console.log");
    knownGlobals.insert("console.error");
    knownGlobals.insert("console.warn");
    knownGlobals.insert("console.info");
    knownGlobals.insert("console.debug");
    knownGlobals.insert("console.trace");
    knownGlobals.insert("console.dir");
    knownGlobals.insert("console.time");
    knownGlobals.insert("console.timeEnd");
    knownGlobals.insert("print");
    
    // Timer Functions (common in JS environments)
    knownGlobals.insert("setTimeout");
    knownGlobals.insert("setInterval");
    knownGlobals.insert("clearTimeout");
    knownGlobals.insert("clearInterval");
    
    // Common Method Names (to reduce false positives)
    knownGlobals.insert("toString");
    knownGlobals.insert("valueOf");
    knownGlobals.insert("hasOwnProperty");
    knownGlobals.insert("isPrototypeOf");
    knownGlobals.insert("propertyIsEnumerable");
    knownGlobals.insert("constructor");
    knownGlobals.insert("prototype");
    knownGlobals.insert("length");
    knownGlobals.insert("name");
    knownGlobals.insert("message");
    knownGlobals.insert("stack");
}

} // namespace macro
} // namespace onlyoffice