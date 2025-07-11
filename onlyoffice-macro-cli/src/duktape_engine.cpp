#include "duktape_engine.h"
#include <fstream>
#include <sstream>
#include <iostream>

namespace onlyoffice {
namespace macro {

DuktapeEngine::DuktapeEngine() : ctx(nullptr), initialized(false) {}

DuktapeEngine::~DuktapeEngine() {
    cleanup();
}

bool DuktapeEngine::initialize() {
    if (initialized) {
        return true;
    }
    
    ctx = duk_create_heap_default();
    if (!ctx) {
        lastError = "Failed to create Duktape context";
        return false;
    }
    
    // Setup basic global functions
    duk_push_c_function(ctx, js_print, DUK_VARARGS);
    duk_put_global_string(ctx, "print");
    
    initialized = true;
    
    // Setup OnlyOffice API mock
    setupOnlyOfficeAPI();
    
    return true;
}

void DuktapeEngine::cleanup() {
    if (ctx) {
        duk_destroy_heap(ctx);
        ctx = nullptr;
    }
    initialized = false;
}

bool DuktapeEngine::executeScript(const std::string& script) {
    if (!initialized) {
        lastError = "Engine not initialized";
        return false;
    }
    
    return loadScript(script);
}

bool DuktapeEngine::executeScriptWithContext(const std::string& script, const std::string& source) {
    if (!initialized) {
        lastError = "Engine not initialized";
        return false;
    }
    
    currentSource = source;
    return loadScript(script);
}

bool DuktapeEngine::executeFile(const std::string& filepath) {
    if (!initialized) {
        lastError = "Engine not initialized";
        return false;
    }
    
    std::string content = readFile(filepath);
    if (content.empty()) {
        lastError = "Failed to read file: " + filepath;
        return false;
    }
    
    return loadScript(content);
}

void DuktapeEngine::setGlobalObject(const std::string& name, duk_c_function func) {
    if (!initialized) return;
    
    duk_push_c_function(ctx, func, DUK_VARARGS);
    duk_put_global_string(ctx, name.c_str());
}

void DuktapeEngine::setGlobalString(const std::string& name, const std::string& value) {
    if (!initialized) return;
    
    duk_push_string(ctx, value.c_str());
    duk_put_global_string(ctx, name.c_str());
}

void DuktapeEngine::setGlobalNumber(const std::string& name, double value) {
    if (!initialized) return;
    
    duk_push_number(ctx, value);
    duk_put_global_string(ctx, name.c_str());
}

void DuktapeEngine::setGlobalBoolean(const std::string& name, bool value) {
    if (!initialized) return;
    
    duk_push_boolean(ctx, value);
    duk_put_global_string(ctx, name.c_str());
}

std::string DuktapeEngine::getLastError() const {
    return lastError;
}

bool DuktapeEngine::hasError() const {
    return !lastError.empty();
}

void DuktapeEngine::setupOnlyOfficeAPI() {
    if (!initialized) {
        return;
    }
    
    pushOnlyOfficeAPI();
}

bool DuktapeEngine::loadScript(const std::string& script) {
    lastError.clear();
    
    duk_push_string(ctx, script.c_str());
    
    if (duk_peval(ctx) != 0) {
        // Get detailed error information
        if (duk_is_error(ctx, -1)) {
            // Try to get line number
            duk_get_prop_string(ctx, -1, "lineNumber");
            if (duk_is_number(ctx, -1)) {
                int line = duk_get_int(ctx, -1);
                duk_pop(ctx);
                
                // Get error message
                std::string message = duk_safe_to_string(ctx, -1);
                
                // Add code context if available
                std::string context;
                if (!currentSource.empty()) {
                    context = getErrorContext(currentSource, line);
                    if (!context.empty()) {
                        // Trim whitespace from context
                        size_t start = context.find_first_not_of(" \t");
                        if (start != std::string::npos) {
                            context = context.substr(start);
                        }
                    }
                }
                
                // Format error with line number and context
                lastError = "Line " + std::to_string(line) + ": " + message;
                if (!context.empty()) {
                    lastError += "\n  Code: " + context;
                }
            } else {
                duk_pop(ctx);
                lastError = duk_safe_to_string(ctx, -1);
            }
        } else {
            lastError = duk_safe_to_string(ctx, -1);
        }
        
        duk_pop(ctx);
        return false;
    }
    
    duk_pop(ctx);
    return true;
}

std::string DuktapeEngine::readFile(const std::string& filepath) {
    std::ifstream file(filepath);
    if (!file.is_open()) {
        return "";
    }
    
    std::ostringstream buffer;
    buffer << file.rdbuf();
    return buffer.str();
}

std::string DuktapeEngine::getErrorContext(const std::string& source, int lineNumber) {
    std::istringstream stream(source);
    std::string line;
    int currentLine = 1;
    
    // Find the problematic line
    while (std::getline(stream, line) && currentLine <= lineNumber) {
        if (currentLine == lineNumber) {
            return line;
        }
        currentLine++;
    }
    
    return "";
}

void DuktapeEngine::pushOnlyOfficeAPI() {
    // Create Api object
    duk_push_object(ctx);
    
    // Add existing methods
    duk_push_c_function(ctx, js_get_active_sheet, 0);
    duk_put_prop_string(ctx, -2, "GetActiveSheet");
    
    duk_push_c_function(ctx, js_show_message, 2);
    duk_put_prop_string(ctx, -2, "ShowMessage");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "GetActiveDocument");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "GetActivePresentation");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "CreateDocument");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "GetDocument");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "GetSheet");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "GetRange");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "GetSelection");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "GetWorkbook");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "GetWorksheet");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "CreateParagraph");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "CreateRun");
    
    duk_push_c_function(ctx, js_api_mock, 0);
    duk_put_prop_string(ctx, -2, "CreateSlide");
    
    // Add the 57 missing API methods
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "AddComment");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "AddCustomFunction");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "AddCustomFunctionLibrary");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "AddDefName");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "AddSheet");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "ClearCustomFunctions");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateBlipFill");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateBullet");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateColorByName");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateColorFromRGB");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateGradientStop");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateLinearGradientFill");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateNewHistoryPoint");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateNoFill");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateNumbering");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreatePatternFill");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreatePresetColor");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateRGBColor");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateRadialGradientFill");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateSchemeColor");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateSolidFill");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateStroke");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "CreateTextPr");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "Format");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetAllComments");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetAllPivotTables");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetCommentById");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetComments");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetDefName");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetDocumentInfo");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetFreezePanesType");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetFullName");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetLocale");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetMailMergeData");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetPivotByName");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetReferenceStyle");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetSheets");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetThemesColors");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "GetWorksheetFunction");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "InsertPivotExistingWorksheet");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "InsertPivotNewWorksheet");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "Intersect");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "OnDocumentReady");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "RecalculateAllFormulas");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "RefreshAllPivots");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "RemoveCustomFunction");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "ReplaceTextSmart");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "Save");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "SetFreezePanesType");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "SetLocale");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "SetReferenceStyle");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "SetThemeColors");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "attachEvent");
    
    duk_push_c_function(ctx, js_api_mock, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "detachEvent");
    
    // Put Api object in global scope
    duk_put_global_string(ctx, "Api");
    
    // Add console object
    duk_push_object(ctx);
    duk_push_c_function(ctx, js_print, DUK_VARARGS);
    duk_put_prop_string(ctx, -2, "log");
    duk_put_global_string(ctx, "console");
}

// Static callback functions
duk_ret_t DuktapeEngine::js_print(duk_context* ctx) {
    duk_idx_t n = duk_get_top(ctx);
    for (duk_idx_t i = 0; i < n; i++) {
        if (i > 0) {
            std::cout << " ";
        }
        std::cout << duk_to_string(ctx, i);
    }
    std::cout << std::endl;
    return 0;
}

duk_ret_t DuktapeEngine::js_api_mock(duk_context* ctx) {
    // Return a mock object
    duk_push_object(ctx);
    return 1;
}

duk_ret_t DuktapeEngine::js_get_active_sheet(duk_context* ctx) {
    // Create a mock sheet object
    duk_push_object(ctx);
    
    // Add GetRange method
    duk_push_c_function(ctx, js_get_range, 1);
    duk_put_prop_string(ctx, -2, "GetRange");
    
    return 1;
}

duk_ret_t DuktapeEngine::js_get_range(duk_context* ctx) {
    // const char* range = duk_require_string(ctx, 0); // unused for now
    
    // Create mock range object
    duk_push_object(ctx);
    
    // Add SetFillColor method
    duk_push_c_function(ctx, js_set_fill_color, 3);
    duk_put_prop_string(ctx, -2, "SetFillColor");
    
    return 1;
}

duk_ret_t DuktapeEngine::js_set_fill_color(duk_context* ctx) {
    int r = duk_require_int(ctx, 0);
    int g = duk_require_int(ctx, 1);  
    int b = duk_require_int(ctx, 2);
    
    std::cout << "Mock: Setting fill color to RGB(" << r << ", " << g << ", " << b << ")" << std::endl;
    return 0;
}

duk_ret_t DuktapeEngine::js_show_message(duk_context* ctx) {
    const char* title = duk_require_string(ctx, 0);
    const char* message = duk_require_string(ctx, 1);
    
    std::cout << "=== " << title << " ===" << std::endl;
    std::cout << message << std::endl;
    std::cout << "===============" << std::endl;
    
    return 0;
}

} // namespace macro
} // namespace onlyoffice