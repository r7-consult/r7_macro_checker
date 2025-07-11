#pragma once

#include "js_engine_interface.h"

#ifdef HAVE_V8
#include <v8.h>
#include <libplatform/libplatform.h>
#endif

namespace onlyoffice {
namespace macro {

#ifdef HAVE_V8

/**
 * @brief V8 JavaScript Engine Implementation
 * 
 * Implements JavaScript engine interface using Google's V8 engine.
 */
class V8Engine : public JSEngineInterface {
public:
    V8Engine();
    ~V8Engine() override;
    
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
    v8::Isolate* isolate_;
    v8::Global<v8::Context> context_;
    std::unique_ptr<v8::Platform> platform_;
    std::string lastError_;
    bool initialized_;
    
    // Helper methods
    void extractV8Error(const v8::TryCatch& try_catch, std::vector<LintIssue>& issues);
    void extractV8Error(const v8::TryCatch& try_catch);
    void setupParameters(v8::Local<v8::Context> context, const std::map<std::string, std::string>& params);
    void setupApiMethods(v8::Local<v8::Object> api);
    void setupConsoleObject(v8::Local<v8::Context> context);
    
    // Static helper methods for API mocking
    static void setupSheetMethods(v8::Local<v8::Object> sheet, v8::Isolate* isolate);
    static void setupRangeMethods(v8::Local<v8::Object> range, v8::Isolate* isolate);
    static void setupDocumentMethods(v8::Local<v8::Object> document, v8::Isolate* isolate);
    static void setupPresentationMethods(v8::Local<v8::Object> presentation, v8::Isolate* isolate);
    
    // V8 callback functions
    static void ApiGetActiveSheet(const v8::FunctionCallbackInfo<v8::Value>& args);
    static void ApiGetActiveDocument(const v8::FunctionCallbackInfo<v8::Value>& args);
    static void ApiGetActivePresentation(const v8::FunctionCallbackInfo<v8::Value>& args);
    static void ApiShowMessage(const v8::FunctionCallbackInfo<v8::Value>& args);
    static void ApiSave(const v8::FunctionCallbackInfo<v8::Value>& args);
    static void ConsoleLog(const v8::FunctionCallbackInfo<v8::Value>& args);
    static void RangeSetValue(const v8::FunctionCallbackInfo<v8::Value>& args);
    static void RangeGetValue(const v8::FunctionCallbackInfo<v8::Value>& args);
    static void SheetGetRange(const v8::FunctionCallbackInfo<v8::Value>& args);
};

#endif // HAVE_V8

} // namespace macro
} // namespace onlyoffice