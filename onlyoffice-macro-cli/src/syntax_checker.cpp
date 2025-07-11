#include "syntax_checker.h"
#include <fstream>
#include <sstream>
#include <regex>
#include <set>

extern "C" {
#include "duktape.h"
}

namespace onlyoffice {
namespace macro {

class SyntaxChecker::Impl {
public:
    bool strictMode = false;
    bool onlyOfficeAPIChecks = true;
    std::set<std::string> knownAPIs;
    
    Impl() {
        setupKnownAPIs();
    }
    
    void setupKnownAPIs() {
        // OnlyOffice API objects and methods
        knownAPIs.insert("Api");
        knownAPIs.insert("Api.GetActiveSheet");
        knownAPIs.insert("Api.GetActiveDocument");
        knownAPIs.insert("Api.GetActivePresentation");
        knownAPIs.insert("Api.ShowMessage");
        knownAPIs.insert("Api.GetDocument");
        knownAPIs.insert("Api.CreateDocument");
        knownAPIs.insert("Api.GetSheet");
        knownAPIs.insert("Api.GetRange");
        knownAPIs.insert("Api.GetSelection");
        
        // Complete Cell API methods - Api Class
        // Document Creation & Management
        knownAPIs.insert("Api.AddComment");
        knownAPIs.insert("Api.AddCustomFunction");
        knownAPIs.insert("Api.AddDefName");
        knownAPIs.insert("Api.AddSheet");
        knownAPIs.insert("Api.ClearCustomFunctions");
        knownAPIs.insert("Api.CreateNewHistoryPoint");
        knownAPIs.insert("Api.Save");
        
        // Color & Fill Creation
        knownAPIs.insert("Api.CreateBlipFill");
        knownAPIs.insert("Api.CreateColorByName");
        knownAPIs.insert("Api.CreateColorFromRGB");
        knownAPIs.insert("Api.CreateGradientStop");
        knownAPIs.insert("Api.CreateLinearGradientFill");
        knownAPIs.insert("Api.CreateNoFill");
        knownAPIs.insert("Api.CreatePatternFill");
        knownAPIs.insert("Api.CreatePresetColor");
        knownAPIs.insert("Api.CreateRadialGradientFill");
        knownAPIs.insert("Api.CreateRGBColor");
        knownAPIs.insert("Api.CreateSchemeColor");
        knownAPIs.insert("Api.CreateSolidFill");
        knownAPIs.insert("Api.CreateStroke");
        
        // Text & Paragraph Creation
        knownAPIs.insert("Api.CreateBullet");
        knownAPIs.insert("Api.CreateNumbering");
        knownAPIs.insert("Api.CreateParagraph");
        knownAPIs.insert("Api.CreateRun");
        knownAPIs.insert("Api.CreateTextPr");
        
        // Data Retrieval & Access
        knownAPIs.insert("Api.GetAllComments");
        knownAPIs.insert("Api.GetAllPivotTables");
        knownAPIs.insert("Api.GetCommentById");
        knownAPIs.insert("Api.GetComments");
        knownAPIs.insert("Api.GetCore");
        knownAPIs.insert("Api.GetCustomProperties");
        knownAPIs.insert("Api.GetDefName");
        knownAPIs.insert("Api.GetDocumentInfo");
        knownAPIs.insert("Api.GetFreezePanesType");
        knownAPIs.insert("Api.GetFullName");
        knownAPIs.insert("Api.GetLocale");
        knownAPIs.insert("Api.GetMailMergeData");
        knownAPIs.insert("Api.GetPivotByName");
        knownAPIs.insert("Api.GetReferenceStyle");
        knownAPIs.insert("Api.GetSheets");
        knownAPIs.insert("Api.GetThemesColors");
        knownAPIs.insert("Api.GetWorksheetFunction");
        
        // Data Manipulation
        knownAPIs.insert("Api.Format");
        knownAPIs.insert("Api.InsertPivotExistingWorksheet");
        knownAPIs.insert("Api.InsertPivotNewWorksheet");
        knownAPIs.insert("Api.Intersect");
        knownAPIs.insert("Api.RecalculateAllFormulas");
        knownAPIs.insert("Api.RefreshAllPivots");
        knownAPIs.insert("Api.RemoveCustomFunction");
        knownAPIs.insert("Api.ReplaceTextSmart");
        
        // Configuration & Settings
        knownAPIs.insert("Api.SetFreezePanesType");
        knownAPIs.insert("Api.SetLocale");
        knownAPIs.insert("Api.SetReferenceStyle");
        knownAPIs.insert("Api.SetThemeColors");
        
        // Event Handling
        knownAPIs.insert("Api.attachEvent");
        knownAPIs.insert("Api.detachEvent");
        knownAPIs.insert("Api.onWorksheetChange");
        
        // ApiRange Class Methods
        // Cell Value Operations
        knownAPIs.insert("ApiRange.GetValue");
        knownAPIs.insert("ApiRange.GetValue2");
        knownAPIs.insert("ApiRange.SetValue");
        knownAPIs.insert("ApiRange.GetText");
        knownAPIs.insert("ApiRange.GetFormula");
        knownAPIs.insert("ApiRange.SetFormulaArray");
        
        // Range Properties
        knownAPIs.insert("ApiRange.GetAddress");
        knownAPIs.insert("ApiRange.GetCells");
        knownAPIs.insert("ApiRange.GetCount");
        knownAPIs.insert("ApiRange.GetCol");
        knownAPIs.insert("ApiRange.GetRow");
        knownAPIs.insert("ApiRange.GetCols");
        knownAPIs.insert("ApiRange.GetRows");
        knownAPIs.insert("ApiRange.GetClassType");
        knownAPIs.insert("ApiRange.GetWorksheet");
        
        // Formatting & Appearance
        knownAPIs.insert("ApiRange.SetFillColor");
        knownAPIs.insert("ApiRange.SetFontColor");
        knownAPIs.insert("ApiRange.SetFontName");
        knownAPIs.insert("ApiRange.SetFontSize");
        knownAPIs.insert("ApiRange.SetBold");
        knownAPIs.insert("ApiRange.SetItalic");
        knownAPIs.insert("ApiRange.SetStrikeout");
        knownAPIs.insert("ApiRange.SetUnderline");
        knownAPIs.insert("ApiRange.SetBorders");
        knownAPIs.insert("ApiRange.SetAlignHorizontal");
        knownAPIs.insert("ApiRange.SetAlignVertical");
        knownAPIs.insert("ApiRange.SetWrap");
        knownAPIs.insert("ApiRange.SetOrientation");
        knownAPIs.insert("ApiRange.SetNumberFormat");
        
        // Range Operations
        knownAPIs.insert("ApiRange.AutoFit");
        knownAPIs.insert("ApiRange.Clear");
        knownAPIs.insert("ApiRange.Copy");
        knownAPIs.insert("ApiRange.Cut");
        knownAPIs.insert("ApiRange.Paste");
        knownAPIs.insert("ApiRange.PasteSpecial");
        knownAPIs.insert("ApiRange.Delete");
        knownAPIs.insert("ApiRange.Insert");
        knownAPIs.insert("ApiRange.Merge");
        knownAPIs.insert("ApiRange.UnMerge");
        knownAPIs.insert("ApiRange.Select");
        
        // Size & Layout
        knownAPIs.insert("ApiRange.SetColumnWidth");
        knownAPIs.insert("ApiRange.SetRowHeight");
        knownAPIs.insert("ApiRange.GetColumnWidth");
        knownAPIs.insert("ApiRange.GetRowHeight");
        knownAPIs.insert("ApiRange.SetHidden");
        knownAPIs.insert("ApiRange.GetHidden");
        
        // Advanced Operations
        knownAPIs.insert("ApiRange.ForEach");
        knownAPIs.insert("ApiRange.Find");
        knownAPIs.insert("ApiRange.FindNext");
        knownAPIs.insert("ApiRange.FindPrevious");
        knownAPIs.insert("ApiRange.Replace");
        knownAPIs.insert("ApiRange.SetAutoFilter");
        knownAPIs.insert("ApiRange.SetSort");
        knownAPIs.insert("ApiRange.SetOffset");
        
        // Comments & Data
        knownAPIs.insert("ApiRange.AddComment");
        knownAPIs.insert("ApiRange.GetComment");
        knownAPIs.insert("ApiRange.GetCharacters");
        knownAPIs.insert("ApiRange.GetAreas");
        knownAPIs.insert("ApiRange.GetDefName");
        knownAPIs.insert("ApiRange.GetPivotTable");
        knownAPIs.insert("ApiRange.End");
        
        // ApiWorksheet Class Methods
        // Sheet Management
        knownAPIs.insert("ApiWorksheet.GetName");
        knownAPIs.insert("ApiWorksheet.SetName");
        knownAPIs.insert("ApiWorksheet.GetIndex");
        knownAPIs.insert("ApiWorksheet.SetActive");
        knownAPIs.insert("ApiWorksheet.GetVisible");
        knownAPIs.insert("ApiWorksheet.SetVisible");
        knownAPIs.insert("ApiWorksheet.Delete");
        knownAPIs.insert("ApiWorksheet.Move");
        
        // Range Operations
        knownAPIs.insert("ApiWorksheet.GetRange");
        knownAPIs.insert("ApiWorksheet.GetRangeByNumber");
        knownAPIs.insert("ApiWorksheet.GetCells");
        knownAPIs.insert("ApiWorksheet.GetUsedRange");
        knownAPIs.insert("ApiWorksheet.GetSelection");
        knownAPIs.insert("ApiWorksheet.GetActiveCell");
        knownAPIs.insert("ApiWorksheet.GetCols");
        knownAPIs.insert("ApiWorksheet.GetRows");
        
        // Objects & Content
        knownAPIs.insert("ApiWorksheet.AddChart");
        knownAPIs.insert("ApiWorksheet.AddShape");
        knownAPIs.insert("ApiWorksheet.AddImage");
        knownAPIs.insert("ApiWorksheet.AddOleObject");
        knownAPIs.insert("ApiWorksheet.AddWordArt");
        knownAPIs.insert("ApiWorksheet.GetAllCharts");
        knownAPIs.insert("ApiWorksheet.GetAllShapes");
        knownAPIs.insert("ApiWorksheet.GetAllImages");
        knownAPIs.insert("ApiWorksheet.GetAllOleObjects");
        knownAPIs.insert("ApiWorksheet.GetAllDrawings");
        
        // Data Management
        knownAPIs.insert("ApiWorksheet.GetDefName");
        knownAPIs.insert("ApiWorksheet.GetDefNames");
        knownAPIs.insert("ApiWorksheet.AddDefName");
        knownAPIs.insert("ApiWorksheet.GetComments");
        knownAPIs.insert("ApiWorksheet.GetAllPivotTables");
        knownAPIs.insert("ApiWorksheet.GetPivotByName");
        knownAPIs.insert("ApiWorksheet.RefreshAllPivots");
        knownAPIs.insert("ApiWorksheet.FormatAsTable");
        
        // Layout & Formatting
        knownAPIs.insert("ApiWorksheet.SetColumnWidth");
        knownAPIs.insert("ApiWorksheet.SetRowHeight");
        knownAPIs.insert("ApiWorksheet.SetDisplayGridlines");
        knownAPIs.insert("ApiWorksheet.SetDisplayHeadings");
        knownAPIs.insert("ApiWorksheet.SetPrintGridlines");
        knownAPIs.insert("ApiWorksheet.SetPrintHeadings");
        knownAPIs.insert("ApiWorksheet.GetPrintGridlines");
        knownAPIs.insert("ApiWorksheet.GetPrintHeadings");
        
        // Page Setup
        knownAPIs.insert("ApiWorksheet.SetPageOrientation");
        knownAPIs.insert("ApiWorksheet.GetPageOrientation");
        knownAPIs.insert("ApiWorksheet.SetLeftMargin");
        knownAPIs.insert("ApiWorksheet.GetLeftMargin");
        knownAPIs.insert("ApiWorksheet.SetRightMargin");
        knownAPIs.insert("ApiWorksheet.GetRightMargin");
        knownAPIs.insert("ApiWorksheet.SetTopMargin");
        knownAPIs.insert("ApiWorksheet.GetTopMargin");
        knownAPIs.insert("ApiWorksheet.SetBottomMargin");
        knownAPIs.insert("ApiWorksheet.GetBottomMargin");
        
        // Freeze Panes & Protection
        knownAPIs.insert("ApiWorksheet.GetFreezePanes");
        knownAPIs.insert("ApiWorksheet.AddProtectedRange");
        knownAPIs.insert("ApiWorksheet.GetProtectedRange");
        knownAPIs.insert("ApiWorksheet.GetAllProtectedRanges");
        
        // Miscellaneous
        knownAPIs.insert("ApiWorksheet.SetHyperlink");
        knownAPIs.insert("ApiWorksheet.Paste");
        knownAPIs.insert("ApiWorksheet.ReplaceCurrentImage");
        
        // ApiChart Class Methods
        // Chart Configuration
        knownAPIs.insert("ApiChart.SetTitle");
        knownAPIs.insert("ApiChart.SetTitleFill");
        knownAPIs.insert("ApiChart.SetTitleOutLine");
        knownAPIs.insert("ApiChart.ApplyChartStyle");
        knownAPIs.insert("ApiChart.GetClassType");
        
        // Series Management
        knownAPIs.insert("ApiChart.AddSeria");
        knownAPIs.insert("ApiChart.RemoveSeria");
        knownAPIs.insert("ApiChart.GetSeries");
        knownAPIs.insert("ApiChart.GetAllSeries");
        knownAPIs.insert("ApiChart.SetSeriaName");
        knownAPIs.insert("ApiChart.SetSeriaValues");
        knownAPIs.insert("ApiChart.SetSeriaXValues");
        knownAPIs.insert("ApiChart.SetSeriesFill");
        knownAPIs.insert("ApiChart.SetSeriesOutLine");
        
        // Data Labels & Markers
        knownAPIs.insert("ApiChart.SetShowDataLabels");
        knownAPIs.insert("ApiChart.SetShowPointDataLabel");
        knownAPIs.insert("ApiChart.SetMarkerFill");
        knownAPIs.insert("ApiChart.SetMarkerOutLine");
        knownAPIs.insert("ApiChart.SetDataPointFill");
        knownAPIs.insert("ApiChart.SetDataPointOutLine");
        
        // Axes Configuration
        knownAPIs.insert("ApiChart.SetHorAxisTitle");
        knownAPIs.insert("ApiChart.SetVerAxisTitle");
        knownAPIs.insert("ApiChart.SetHorAxisOrientation");
        knownAPIs.insert("ApiChart.SetVerAxisOrientation");
        knownAPIs.insert("ApiChart.SetHorAxisMajorTickMark");
        knownAPIs.insert("ApiChart.SetHorAxisMinorTickMark");
        knownAPIs.insert("ApiChart.SetVertAxisMajorTickMark");
        knownAPIs.insert("ApiChart.SetVertAxisMinorTickMark");
        knownAPIs.insert("ApiChart.SetHorAxisTickLabelPosition");
        knownAPIs.insert("ApiChart.SetVertAxisTickLabelPosition");
        knownAPIs.insert("ApiChart.SetHorAxisLablesFontSize");
        knownAPIs.insert("ApiChart.SetVertAxisLablesFontSize");
        knownAPIs.insert("ApiChart.SetAxieNumFormat");
        
        // Gridlines
        knownAPIs.insert("ApiChart.SetMajorHorizontalGridlines");
        knownAPIs.insert("ApiChart.SetMinorHorizontalGridlines");
        knownAPIs.insert("ApiChart.SetMajorVerticalGridlines");
        knownAPIs.insert("ApiChart.SetMinorVerticalGridlines");
        
        // Legend & Plot Area
        knownAPIs.insert("ApiChart.SetLegendPos");
        knownAPIs.insert("ApiChart.SetLegendFill");
        knownAPIs.insert("ApiChart.SetLegendOutLine");
        knownAPIs.insert("ApiChart.SetLegendFontSize");
        knownAPIs.insert("ApiChart.SetPlotAreaFill");
        knownAPIs.insert("ApiChart.SetPlotAreaOutLine");
        
        // Data Source
        knownAPIs.insert("ApiChart.SetCatFormula");
        
        // ApiPivotTable Class Methods
        // Basic Properties
        knownAPIs.insert("ApiPivotTable.GetName");
        knownAPIs.insert("ApiPivotTable.SetName");
        knownAPIs.insert("ApiPivotTable.GetDescription");
        knownAPIs.insert("ApiPivotTable.SetDescription");
        knownAPIs.insert("ApiPivotTable.GetTitle");
        knownAPIs.insert("ApiPivotTable.SetTitle");
        knownAPIs.insert("ApiPivotTable.GetParent");
        
        // Field Management
        knownAPIs.insert("ApiPivotTable.AddFields");
        knownAPIs.insert("ApiPivotTable.AddDataField");
        knownAPIs.insert("ApiPivotTable.RemoveField");
        knownAPIs.insert("ApiPivotTable.MoveField");
        knownAPIs.insert("ApiPivotTable.GetPivotFields");
        knownAPIs.insert("ApiPivotTable.GetColumnFields");
        knownAPIs.insert("ApiPivotTable.GetRowFields");
        knownAPIs.insert("ApiPivotTable.GetPageFields");
        knownAPIs.insert("ApiPivotTable.GetDataFields");
        knownAPIs.insert("ApiPivotTable.GetHiddenFields");
        knownAPIs.insert("ApiPivotTable.GetVisibleFields");
        
        // Data Operations
        knownAPIs.insert("ApiPivotTable.GetData");
        knownAPIs.insert("ApiPivotTable.GetPivotData");
        knownAPIs.insert("ApiPivotTable.RefreshTable");
        knownAPIs.insert("ApiPivotTable.Update");
        knownAPIs.insert("ApiPivotTable.ClearTable");
        knownAPIs.insert("ApiPivotTable.ClearAllFilters");
        
        // Layout & Formatting
        knownAPIs.insert("ApiPivotTable.SetRowAxisLayout");
        knownAPIs.insert("ApiPivotTable.SetLayoutBlankLine");
        knownAPIs.insert("ApiPivotTable.SetLayoutSubtotals");
        knownAPIs.insert("ApiPivotTable.SetSubtotalLocation");
        knownAPIs.insert("ApiPivotTable.SetRepeatAllLabels");
        
        // Grand Totals
        knownAPIs.insert("ApiPivotTable.GetColumnGrand");
        knownAPIs.insert("ApiPivotTable.SetColumnGrand");
        knownAPIs.insert("ApiPivotTable.GetRowGrand");
        knownAPIs.insert("ApiPivotTable.SetRowGrand");
        knownAPIs.insert("ApiPivotTable.GetGrandTotalName");
        knownAPIs.insert("ApiPivotTable.SetGrandTotalName");
        
        // Style & Appearance
        knownAPIs.insert("ApiPivotTable.GetStyleName");
        knownAPIs.insert("ApiPivotTable.SetStyleName");
        knownAPIs.insert("ApiPivotTable.GetTableStyleColumnHeaders");
        knownAPIs.insert("ApiPivotTable.SetTableStyleColumnHeaders");
        knownAPIs.insert("ApiPivotTable.GetTableStyleRowHeaders");
        knownAPIs.insert("ApiPivotTable.SetTableStyleRowHeaders");
        knownAPIs.insert("ApiPivotTable.GetTableStyleColumnStripes");
        knownAPIs.insert("ApiPivotTable.SetTableStyleColumnStripes");
        knownAPIs.insert("ApiPivotTable.GetTableStyleRowStripes");
        knownAPIs.insert("ApiPivotTable.SetTableStyleRowStripes");
        
        // Display Options
        knownAPIs.insert("ApiPivotTable.GetDisplayFieldCaptions");
        knownAPIs.insert("ApiPivotTable.SetDisplayFieldCaptions");
        knownAPIs.insert("ApiPivotTable.GetDisplayFieldsInReportFilterArea");
        knownAPIs.insert("ApiPivotTable.SetDisplayFieldsInReportFilterArea");
        
        // Range Operations
        knownAPIs.insert("ApiPivotTable.GetTableRange1");
        knownAPIs.insert("ApiPivotTable.GetTableRange2");
        knownAPIs.insert("ApiPivotTable.GetColumnRange");
        knownAPIs.insert("ApiPivotTable.GetRowRange");
        knownAPIs.insert("ApiPivotTable.GetDataBodyRange");
        knownAPIs.insert("ApiPivotTable.GetSource");
        knownAPIs.insert("ApiPivotTable.SetSource");
        
        // Interaction
        knownAPIs.insert("ApiPivotTable.Select");
        knownAPIs.insert("ApiPivotTable.ShowDetails");
        knownAPIs.insert("ApiPivotTable.PivotValueCell");
        
        // Common JavaScript objects that should be available
        knownAPIs.insert("console");
        knownAPIs.insert("console.log");
        knownAPIs.insert("JSON");
        knownAPIs.insert("JSON.parse");
        knownAPIs.insert("JSON.stringify");
        knownAPIs.insert("Math");
        knownAPIs.insert("Date");
        knownAPIs.insert("String");
        knownAPIs.insert("Number");
        knownAPIs.insert("Boolean");
        knownAPIs.insert("Array");
        knownAPIs.insert("Object");
        knownAPIs.insert("RegExp");
    }
};

SyntaxChecker::SyntaxChecker() : pImpl(std::make_unique<Impl>()) {}

SyntaxChecker::~SyntaxChecker() = default;

SyntaxCheckResult SyntaxChecker::checkFile(const std::string& filepath) {
    std::ifstream file(filepath);
    if (!file.is_open()) {
        SyntaxCheckResult result;
        result.isValid = false;
        SyntaxError error = {0, 0, "Cannot open file: " + filepath, "error"};
        result.errors.push_back(error);
        return result;
    }
    
    std::ostringstream buffer;
    buffer << file.rdbuf();
    std::string source = buffer.str();
    
    return checkString(source);
}

SyntaxCheckResult SyntaxChecker::checkString(const std::string& source) {
    SyntaxCheckResult result;
    result.source = source;
    result.isValid = true;
    
    // Check JavaScript syntax using Duktape
    if (!checkJavaScriptSyntax(source, result.errors)) {
        result.isValid = false;
    }
    
    // Check OnlyOffice API usage
    if (pImpl->onlyOfficeAPIChecks) {
        if (!validateOnlyOfficeAPI(source, result.errors)) {
            // API validation warnings don't make the syntax invalid
        }
    }
    
    return result;
}

void SyntaxChecker::setStrictMode(bool strict) {
    pImpl->strictMode = strict;
}

void SyntaxChecker::setOnlyOfficeAPIChecks(bool enable) {
    pImpl->onlyOfficeAPIChecks = enable;
}

bool SyntaxChecker::checkJavaScriptSyntax(const std::string& source, std::vector<SyntaxError>& errors) {
    duk_context* ctx = duk_create_heap_default();
    if (!ctx) {
        SyntaxError error = {0, 0, "Failed to create JavaScript context", "error"};
        errors.push_back(error);
        return false;
    }
    
    // Try to compile the script
    duk_push_string(ctx, source.c_str());
    duk_push_string(ctx, "syntax-check");
    
    bool isValid = true;
    
    if (duk_pcompile(ctx, 0) != 0) {
        isValid = false;
        
        // Get error information
        if (duk_is_error(ctx, -1)) {
            duk_get_prop_string(ctx, -1, "lineNumber");
            int line = duk_get_int_default(ctx, -1, 0);
            duk_pop(ctx);
            
            duk_get_prop_string(ctx, -1, "columnNumber");
            int column = duk_get_int_default(ctx, -1, 0);
            duk_pop(ctx);
            
            std::string message = duk_safe_to_string(ctx, -1);
            SyntaxError error = {line, column, message, "error"};
            errors.push_back(error);
        } else {
            std::string message = duk_safe_to_string(ctx, -1);
            SyntaxError error = {0, 0, message, "error"};
            errors.push_back(error);
        }
    }
    
    duk_destroy_heap(ctx);
    return isValid;
}

bool SyntaxChecker::validateOnlyOfficeAPI(const std::string& source, std::vector<SyntaxError>& errors) {
    // Simple regex-based API validation
    std::regex apiCallRegex(R"((\w+(?:\.\w+)*)\s*\()");
    std::sregex_iterator iter(source.begin(), source.end(), apiCallRegex);
    std::sregex_iterator end;
    
    int lineNum = 1;
    size_t lastPos = 0;
    
    for (; iter != end; ++iter) {
        const std::smatch& match = *iter;
        std::string apiCall = match[1].str();
        
        // Count line numbers
        size_t pos = match.position();
        lineNum += std::count(source.begin() + lastPos, source.begin() + pos, '\n');
        lastPos = pos;
        
        // Check if API call is known
        if (apiCall.find("Api.") == 0 || apiCall == "Api") {
            if (pImpl->knownAPIs.find(apiCall) == pImpl->knownAPIs.end()) {
                SyntaxError error;
                error.line = lineNum;
                error.column = static_cast<int>(pos % 80); // Simple column approximation
                error.message = "Unknown OnlyOffice API call: " + apiCall;
                error.severity = "warning";
                errors.push_back(error);
            }
        }
    }
    
    // Check for common mistakes
    if (source.find("document.") != std::string::npos) {
        SyntaxError error = {0, 0, "Use 'Api.GetActiveDocument()' instead of 'document' in OnlyOffice macros", "warning"};
        errors.push_back(error);
    }
    
    if (source.find("window.") != std::string::npos) {
        SyntaxError error = {0, 0, "'window' object is not available in OnlyOffice macros", "warning"};
        errors.push_back(error);
    }
    
    if (source.find("alert(") != std::string::npos) {
        SyntaxError error = {0, 0, "Use 'Api.ShowMessage()' instead of 'alert()' in OnlyOffice macros", "warning"};
        errors.push_back(error);
    }
    
    return true;
}

void SyntaxChecker::addKnownAPIs() {
    // Additional APIs can be added here
    pImpl->knownAPIs.insert("Api.GetWorkbook");
    pImpl->knownAPIs.insert("Api.GetWorksheet");
    pImpl->knownAPIs.insert("Api.CreateParagraph");
    pImpl->knownAPIs.insert("Api.CreateRun");
    pImpl->knownAPIs.insert("Api.CreateSlide");
}

} // namespace macro
} // namespace onlyoffice