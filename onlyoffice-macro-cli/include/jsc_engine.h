#pragma once

#include "js_engine_interface.h"

#ifdef HAVE_JSC
#include <JavaScriptCore/JavaScriptCore.h>
#endif

namespace onlyoffice {
namespace macro {

#ifdef HAVE_JSC

/**
 * @brief JavaScriptCore Engine Implementation
 * 
 * Implements JavaScript engine interface using Apple's JavaScriptCore engine.
 */
class JSCEngine : public JSEngineInterface {
public:
    JSCEngine();
    ~JSCEngine() override;
    
    // JSEngineInterface implementation
    bool initialize() override;
    void cleanup() override;
    bool validateSyntax(const std::string& source, std::vector<LintIssue>& issues) override;
    bool executeScript(const std::string& source, const std::map<std::string, std::string>& params = {}) override;
    void setupOnlyOfficeAPI() override;
    std::string getLastError() const override;
    void clearErrors() override;
    std::string getEngineName() const override;
    std::string getEngineVersion() const override;
    bool isInitialized() const override;

private:
    JSGlobalContextRef context_;
    std::string lastError_;
    bool initialized_;
    
    // Helper methods
    void extractJSCError(JSValueRef exception, std::vector<LintIssue>& issues);
    void extractJSCError(JSValueRef exception);
    int extractLineNumber(JSStringRef errorString);
    int extractColumnNumber(JSStringRef errorString);
    std::string convertJSStringToStdString(JSStringRef jsString);
    void setupParameters(const std::map<std::string, std::string>& params);
    void setupApiMethods(JSObjectRef api);
    void setupConsoleObject();
    
    // Static helper methods for API mocking
    static void setupSheetMethods(JSContextRef ctx, JSObjectRef sheet);
    static void setupRangeMethods(JSContextRef ctx, JSObjectRef range);
    static void setupDocumentMethods(JSContextRef ctx, JSObjectRef document);
    static void setupPresentationMethods(JSContextRef ctx, JSObjectRef presentation);
    
    // JSC callback functions
    static JSValueRef ApiGetActiveSheet(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception);
    static JSValueRef ApiGetActiveDocument(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception);
    static JSValueRef ApiGetActivePresentation(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception);
    static JSValueRef ApiShowMessage(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception);
    static JSValueRef ApiSave(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception);
    static JSValueRef ConsoleLog(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception);
    static JSValueRef RangeSetValue(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception);
    static JSValueRef RangeGetValue(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception);
    static JSValueRef SheetGetRange(JSContextRef ctx, JSObjectRef function, JSObjectRef thisObject, size_t argumentCount, const JSValueRef arguments[], JSValueRef* exception);
    
    // Utility functions
    static std::string convertJSStringToStdString(JSStringRef jsString);
    static JSStringRef convertStdStringToJSString(const std::string& str);
};

#endif // HAVE_JSC

} // namespace macro
} // namespace onlyoffice