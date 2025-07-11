#include "jsc_engine.h"

#ifdef HAVE_JSC

#include <iostream>
#include <sstream>
#include <regex>

namespace onlyoffice {
namespace macro {

JSCEngine::JSCEngine() : context_(nullptr), initialized_(false) {
}

JSCEngine::~JSCEngine() {
    cleanup();
}

bool JSCEngine::initialize() {
    if (initialized_) {
        return true;
    }
    
    try {
        context_ = JSGlobalContextCreate(nullptr);
        if (!context_) {
            lastError_ = "Failed to create JSC context";
            return false;
        }
        
        initialized_ = true;
        return true;
    } catch (const std::exception& e) {
        lastError_ = "JSC initialization failed: " + std::string(e.what());
        return false;
    }
}

void JSCEngine::cleanup() {
    if (!initialized_) {
        return;
    }
    
    if (context_) {
        JSGlobalContextRelease(context_);
        context_ = nullptr;
    }
    
    initialized_ = false;
}

bool JSCEngine::validateSyntax(const std::string& source, std::vector<LintIssue>& issues) {
    if (!initialized_ || !context_) {
        return false;
    }
    
    JSStringRef script = JSStringCreateWithUTF8CString(source.c_str());
    JSValueRef exception = nullptr;
    
    // Check syntax
    bool isValid = JSCheckScriptSyntax(context_, script, nullptr, 0, &exception);
    
    if (!isValid && exception) {
        extractJSCError(exception, issues);
    }
    
    JSStringRelease(script);
    return isValid;
}

bool JSCEngine::executeScript(const std::string& source, const std::map<std::string, std::string>& params) {
    if (!initialized_ || !context_) {
        return false;
    }
    
    // Set up parameters
    if (!params.empty()) {
        setupParameters(params);
    }
    
    JSStringRef script = JSStringCreateWithUTF8CString(source.c_str());
    JSValueRef exception = nullptr;
    
    JSValueRef result = JSEvaluateScript(context_, script, nullptr, nullptr, 0, &exception);
    
    JSStringRelease(script);
    
    if (exception) {
        extractJSCError(exception);
        return false;
    }
    
    return true;
}

void JSCEngine::setupOnlyOfficeAPI() {
    if (!initialized_ || !context_) {
        return;
    }
    
    // Create Api object
    JSObjectRef api = JSObjectMake(context_, nullptr, nullptr);
    
    // Add mock methods
    setupApiMethods(api);
    
    // Set global Api
    JSStringRef apiName = JSStringCreateWithUTF8CString("Api");
    JSObjectSetProperty(context_, JSContextGetGlobalObject(context_), apiName, api, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(apiName);
    
    // Setup console
    setupConsoleObject();
}

std::string JSCEngine::getLastError() const {
    return lastError_;
}

void JSCEngine::clearErrors() {
    lastError_.clear();
}

std::string JSCEngine::getEngineName() const {
    return "JavaScriptCore";
}

std::string JSCEngine::getEngineVersion() const {
    return "System JSC";
}

bool JSCEngine::isInitialized() const {
    return initialized_;
}

void JSCEngine::extractJSCError(JSValueRef exception, std::vector<LintIssue>& issues) {
    JSStringRef exceptionString = JSValueToStringCopy(context_, exception, nullptr);
    
    LintIssue issue;
    issue.line = extractLineNumber(exceptionString);
    issue.column = extractColumnNumber(exceptionString);
    issue.message = convertJSStringToStdString(exceptionString);
    issue.rule = "syntax";
    issue.severity = LintSeverity::Error;
    
    issues.push_back(issue);
    JSStringRelease(exceptionString);
}

void JSCEngine::extractJSCError(JSValueRef exception) {
    JSStringRef exceptionString = JSValueToStringCopy(context_, exception, nullptr);
    lastError_ = convertJSStringToStdString(exceptionString);
    JSStringRelease(exceptionString);
}

int JSCEngine::extractLineNumber(JSStringRef errorString) {
    // Parse error string to extract line number
    // JSC error format: "SyntaxError: message (line X)"
    std::string error = convertJSStringToStdString(errorString);
    
    // Try different patterns for line number extraction
    std::regex linePattern(R"(\(line\s+(\d+)\))");
    std::regex linePattern2(R"(line\s+(\d+))");
    std::regex linePattern3(R"(:(\d+):)");
    
    std::smatch match;
    if (std::regex_search(error, match, linePattern) && match.size() > 1) {
        return std::stoi(match[1].str());
    } else if (std::regex_search(error, match, linePattern2) && match.size() > 1) {
        return std::stoi(match[1].str());
    } else if (std::regex_search(error, match, linePattern3) && match.size() > 1) {
        return std::stoi(match[1].str());
    }
    
    return 0;
}

int JSCEngine::extractColumnNumber(JSStringRef errorString) {
    // JSC doesn't always provide column info in basic error messages
    std::string error = convertJSStringToStdString(errorString);
    
    // Try to extract column from format like ":line:column:"
    std::regex colPattern(R"(:(\d+):(\d+):)");
    std::smatch match;
    if (std::regex_search(error, match, colPattern) && match.size() > 2) {
        return std::stoi(match[2].str());
    }
    
    return 0;
}

std::string JSCEngine::convertJSStringToStdString(JSStringRef jsString) {
    size_t length = JSStringGetMaximumUTF8CStringSize(jsString);
    char* buffer = new char[length];
    JSStringGetUTF8CString(jsString, buffer, length);
    std::string result(buffer);
    delete[] buffer;
    return result;
}

void JSCEngine::setupParameters(const std::map<std::string, std::string>& params) {
    JSObjectRef paramsObj = JSObjectMake(context_, nullptr, nullptr);
    
    for (const auto& param : params) {
        JSStringRef key = JSStringCreateWithUTF8CString(param.first.c_str());
        JSStringRef value = JSStringCreateWithUTF8CString(param.second.c_str());
        JSValueRef valueRef = JSValueMakeString(context_, value);
        
        JSObjectSetProperty(context_, paramsObj, key, valueRef, kJSPropertyAttributeNone, nullptr);
        
        JSStringRelease(key);
        JSStringRelease(value);
    }
    
    JSStringRef paramsName = JSStringCreateWithUTF8CString("PARAMS");
    JSObjectSetProperty(context_, JSContextGetGlobalObject(context_), paramsName, paramsObj, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(paramsName);
}

void JSCEngine::setupApiMethods(JSObjectRef api) {
    // GetActiveSheet mock
    JSStringRef getActiveSheetName = JSStringCreateWithUTF8CString("GetActiveSheet");
    JSObjectRef getActiveSheetFunc = JSObjectMakeFunctionWithCallback(context_, getActiveSheetName, ApiGetActiveSheet);
    JSObjectSetProperty(context_, api, getActiveSheetName, getActiveSheetFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(getActiveSheetName);
    
    // GetActiveDocument mock
    JSStringRef getActiveDocumentName = JSStringCreateWithUTF8CString("GetActiveDocument");
    JSObjectRef getActiveDocumentFunc = JSObjectMakeFunctionWithCallback(context_, getActiveDocumentName, ApiGetActiveDocument);
    JSObjectSetProperty(context_, api, getActiveDocumentName, getActiveDocumentFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(getActiveDocumentName);
    
    // GetActivePresentation mock
    JSStringRef getActivePresentationName = JSStringCreateWithUTF8CString("GetActivePresentation");
    JSObjectRef getActivePresentationFunc = JSObjectMakeFunctionWithCallback(context_, getActivePresentationName, ApiGetActivePresentation);
    JSObjectSetProperty(context_, api, getActivePresentationName, getActivePresentationFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(getActivePresentationName);
    
    // ShowMessage mock
    JSStringRef showMessageName = JSStringCreateWithUTF8CString("ShowMessage");
    JSObjectRef showMessageFunc = JSObjectMakeFunctionWithCallback(context_, showMessageName, ApiShowMessage);
    JSObjectSetProperty(context_, api, showMessageName, showMessageFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(showMessageName);
    
    // Save mock
    JSStringRef saveName = JSStringCreateWithUTF8CString("Save");
    JSObjectRef saveFunc = JSObjectMakeFunctionWithCallback(context_, saveName, ApiSave);
    JSObjectSetProperty(context_, api, saveName, saveFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(saveName);
}

void JSCEngine::setupConsoleObject() {
    JSObjectRef console = JSObjectMake(context_, nullptr, nullptr);
    
    // console.log
    JSStringRef logName = JSStringCreateWithUTF8CString("log");
    JSObjectRef logFunc = JSObjectMakeFunctionWithCallback(context_, logName, ConsoleLog);
    JSObjectSetProperty(context_, console, logName, logFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(logName);
    
    JSStringRef consoleName = JSStringCreateWithUTF8CString("console");
    JSObjectSetProperty(context_, JSContextGetGlobalObject(context_), consoleName, console, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(consoleName);
}

void JSCEngine::setupSheetMethods(JSContextRef ctx, JSObjectRef sheet) {
    JSStringRef getRangeName = JSStringCreateWithUTF8CString("GetRange");
    JSObjectRef getRangeFunc = JSObjectMakeFunctionWithCallback(ctx, getRangeName, SheetGetRange);
    JSObjectSetProperty(ctx, sheet, getRangeName, getRangeFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(getRangeName);
}

void JSCEngine::setupRangeMethods(JSContextRef ctx, JSObjectRef range) {
    JSStringRef setValueName = JSStringCreateWithUTF8CString("SetValue");
    JSObjectRef setValueFunc = JSObjectMakeFunctionWithCallback(ctx, setValueName, RangeSetValue);
    JSObjectSetProperty(ctx, range, setValueName, setValueFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(setValueName);
    
    JSStringRef getValueName = JSStringCreateWithUTF8CString("GetValue");
    JSObjectRef getValueFunc = JSObjectMakeFunctionWithCallback(ctx, getValueName, RangeGetValue);
    JSObjectSetProperty(ctx, range, getValueName, getValueFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(getValueName);
}

void JSCEngine::setupDocumentMethods(JSContextRef ctx, JSObjectRef document) {
    JSStringRef createParagraphName = JSStringCreateWithUTF8CString("CreateParagraph");
    JSObjectRef createParagraphFunc = JSObjectMakeFunctionWithCallback(ctx, createParagraphName, [](JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) -> JSValueRef {
        JSObjectRef paragraph = JSObjectMake(ctx, nullptr, nullptr);
        return paragraph;
    });
    JSObjectSetProperty(ctx, document, createParagraphName, createParagraphFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(createParagraphName);
}

void JSCEngine::setupPresentationMethods(JSContextRef ctx, JSObjectRef presentation) {
    JSStringRef getSlideByIndexName = JSStringCreateWithUTF8CString("GetSlideByIndex");
    JSObjectRef getSlideByIndexFunc = JSObjectMakeFunctionWithCallback(ctx, getSlideByIndexName, [](JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) -> JSValueRef {
        JSObjectRef slide = JSObjectMake(ctx, nullptr, nullptr);
        return slide;
    });
    JSObjectSetProperty(ctx, presentation, getSlideByIndexName, getSlideByIndexFunc, kJSPropertyAttributeNone, nullptr);
    JSStringRelease(getSlideByIndexName);
}

// Static callback implementations
JSValueRef JSCEngine::ApiGetActiveSheet(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) {
    JSObjectRef sheet = JSObjectMake(ctx, nullptr, nullptr);
    setupSheetMethods(ctx, sheet);
    return sheet;
}

JSValueRef JSCEngine::ApiGetActiveDocument(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) {
    JSObjectRef document = JSObjectMake(ctx, nullptr, nullptr);
    setupDocumentMethods(ctx, document);
    return document;
}

JSValueRef JSCEngine::ApiGetActivePresentation(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) {
    JSObjectRef presentation = JSObjectMake(ctx, nullptr, nullptr);
    setupPresentationMethods(ctx, presentation);
    return presentation;
}

JSValueRef JSCEngine::ApiShowMessage(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) {
    if (argumentCount >= 2) {
        JSStringRef title = JSValueToStringCopy(ctx, arguments[0], nullptr);
        JSStringRef message = JSValueToStringCopy(ctx, arguments[1], nullptr);
        
        std::string titleStr = convertJSStringToStdString(title);
        std::string messageStr = convertJSStringToStdString(message);
        
        printf("📋 Api.ShowMessage: %s - %s\n", titleStr.c_str(), messageStr.c_str());
        
        JSStringRelease(title);
        JSStringRelease(message);
    }
    return JSValueMakeUndefined(ctx);
}

JSValueRef JSCEngine::ApiSave(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) {
    printf("💾 Api.Save: Document saved\n");
    return JSValueMakeUndefined(ctx);
}

JSValueRef JSCEngine::ConsoleLog(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) {
    for (size_t i = 0; i < argumentCount; i++) {
        JSStringRef str = JSValueToStringCopy(ctx, arguments[i], nullptr);
        std::string strVal = convertJSStringToStdString(str);
        printf("%s%s", i > 0 ? " " : "", strVal.c_str());
        JSStringRelease(str);
    }
    printf("\n");
    return JSValueMakeUndefined(ctx);
}

JSValueRef JSCEngine::RangeSetValue(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) {
    if (argumentCount >= 1) {
        JSStringRef value = JSValueToStringCopy(ctx, arguments[0], nullptr);
        std::string valueStr = convertJSStringToStdString(value);
        printf("📝 Range.SetValue: %s\n", valueStr.c_str());
        JSStringRelease(value);
    }
    return JSValueMakeUndefined(ctx);
}

JSValueRef JSCEngine::RangeGetValue(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) {
    printf("📖 Range.GetValue: MockValue\n");
    JSStringRef mockValue = JSStringCreateWithUTF8CString("MockValue");
    JSValueRef result = JSValueMakeString(ctx, mockValue);
    JSStringRelease(mockValue);
    return result;
}

JSValueRef JSCEngine::SheetGetRange(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception) {
    JSObjectRef range = JSObjectMake(ctx, nullptr, nullptr);
    setupRangeMethods(ctx, range);
    
    if (argumentCount >= 1) {
        JSStringRef rangeAddr = JSValueToStringCopy(ctx, arguments[0], nullptr);
        std::string rangeAddrStr = convertJSStringToStdString(rangeAddr);
        printf("📊 Sheet.GetRange: %s\n", rangeAddrStr.c_str());
        JSStringRelease(rangeAddr);
    }
    
    return range;
}

// Utility functions
std::string JSCEngine::convertJSStringToStdString(JSStringRef jsString) {
    size_t length = JSStringGetMaximumUTF8CStringSize(jsString);
    char* buffer = new char[length];
    JSStringGetUTF8CString(jsString, buffer, length);
    std::string result(buffer);
    delete[] buffer;
    return result;
}

JSStringRef JSCEngine::convertStdStringToJSString(const std::string& str) {
    return JSStringCreateWithUTF8CString(str.c_str());
}

} // namespace macro
} // namespace onlyoffice

#endif // HAVE_JSC