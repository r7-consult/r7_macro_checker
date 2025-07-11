#include "v8_engine.h"

#ifdef HAVE_V8

#include <iostream>
#include <sstream>

namespace onlyoffice {
namespace macro {

V8Engine::V8Engine() : isolate_(nullptr), initialized_(false) {
}

V8Engine::~V8Engine() {
    cleanup();
}

bool V8Engine::initialize() {
    if (initialized_) {
        return true;
    }
    
    try {
        // Initialize V8 platform
        platform_ = v8::platform::NewDefaultPlatform();
        v8::V8::InitializePlatform(platform_.get());
        v8::V8::Initialize();
        
        // Create isolate
        v8::Isolate::CreateParams create_params;
        create_params.array_buffer_allocator = v8::ArrayBuffer::Allocator::NewDefaultAllocator();
        isolate_ = v8::Isolate::New(create_params);
        
        if (!isolate_) {
            lastError_ = "Failed to create V8 isolate";
            return false;
        }
        
        // Create context
        v8::Isolate::Scope isolate_scope(isolate_);
        v8::HandleScope handle_scope(isolate_);
        v8::Local<v8::Context> context = v8::Context::New(isolate_);
        context_.Reset(isolate_, context);
        
        initialized_ = true;
        return true;
    } catch (const std::exception& e) {
        lastError_ = "V8 initialization failed: " + std::string(e.what());
        return false;
    }
}

void V8Engine::cleanup() {
    if (!initialized_) {
        return;
    }
    
    if (isolate_) {
        context_.Reset();
        isolate_->Dispose();
        isolate_ = nullptr;
    }
    
    if (platform_) {
        v8::V8::Dispose();
        v8::V8::ShutdownPlatform();
        platform_.reset();
    }
    
    initialized_ = false;
}

bool V8Engine::validateSyntax(const std::string& source, std::vector<LintIssue>& issues) {
    if (!initialized_ || !isolate_) {
        return false;
    }
    
    v8::Isolate::Scope isolate_scope(isolate_);
    v8::HandleScope handle_scope(isolate_);
    v8::Local<v8::Context> context = context_.Get(isolate_);
    v8::Context::Scope context_scope(context);
    
    v8::TryCatch try_catch(isolate_);
    
    // Create script source
    v8::Local<v8::String> v8_source = v8::String::NewFromUtf8(isolate_, source.c_str()).ToLocalChecked();
    v8::ScriptOrigin origin(v8::String::NewFromUtf8(isolate_, "macro-validation").ToLocalChecked());
    
    // Try to compile script
    v8::Local<v8::Script> script;
    if (!v8::Script::Compile(context, v8_source, &origin).ToLocal(&script)) {
        // Extract compilation error
        extractV8Error(try_catch, issues);
        return false;
    }
    
    return true;
}

bool V8Engine::executeScript(const std::string& source, const std::map<std::string, std::string>& params) {
    if (!initialized_ || !isolate_) {
        return false;
    }
    
    v8::Isolate::Scope isolate_scope(isolate_);
    v8::HandleScope handle_scope(isolate_);
    v8::Local<v8::Context> context = context_.Get(isolate_);
    v8::Context::Scope context_scope(context);
    
    v8::TryCatch try_catch(isolate_);
    
    // Set up parameters
    if (!params.empty()) {
        setupParameters(context, params);
    }
    
    // Execute script
    v8::Local<v8::String> v8_source = v8::String::NewFromUtf8(isolate_, source.c_str()).ToLocalChecked();
    v8::Local<v8::Script> script;
    
    if (!v8::Script::Compile(context, v8_source).ToLocal(&script)) {
        extractV8Error(try_catch);
        return false;
    }
    
    v8::Local<v8::Value> result;
    if (!script->Run(context).ToLocal(&result)) {
        extractV8Error(try_catch);
        return false;
    }
    
    return true;
}

void V8Engine::setupOnlyOfficeAPI() {
    if (!initialized_ || !isolate_) {
        return;
    }
    
    v8::Isolate::Scope isolate_scope(isolate_);
    v8::HandleScope handle_scope(isolate_);
    v8::Local<v8::Context> context = context_.Get(isolate_);
    v8::Context::Scope context_scope(context);
    
    // Create Api object
    v8::Local<v8::Object> api = v8::Object::New(isolate_);
    
    // Add mock methods
    setupApiMethods(api);
    
    // Set global Api object
    context->Global()->Set(context, 
        v8::String::NewFromUtf8(isolate_, "Api").ToLocalChecked(), 
        api).Check();
    
    // Setup console object
    setupConsoleObject(context);
}

std::string V8Engine::getLastError() const {
    return lastError_;
}

void V8Engine::clearErrors() {
    lastError_.clear();
}

std::string V8Engine::getEngineName() const {
    return "V8";
}

std::string V8Engine::getEngineVersion() const {
    return v8::V8::GetVersion();
}

bool V8Engine::isInitialized() const {
    return initialized_;
}

void V8Engine::extractV8Error(const v8::TryCatch& try_catch, std::vector<LintIssue>& issues) {
    v8::Local<v8::Message> message = try_catch.Message();
    v8::Local<v8::Value> exception = try_catch.Exception();
    
    if (!message.IsEmpty()) {
        LintIssue issue;
        issue.line = message->GetLineNumber(isolate_->GetCurrentContext()).FromMaybe(0);
        issue.column = message->GetStartColumn(isolate_->GetCurrentContext()).FromMaybe(0);
        issue.message = *v8::String::Utf8Value(isolate_, exception);
        issue.rule = "syntax";
        issue.severity = LintSeverity::Error;
        
        // Extract source context
        v8::Local<v8::String> sourceLine = message->GetSourceLine(isolate_->GetCurrentContext()).ToLocalChecked();
        issue.codeContext = *v8::String::Utf8Value(isolate_, sourceLine);
        
        issues.push_back(issue);
    }
}

void V8Engine::extractV8Error(const v8::TryCatch& try_catch) {
    v8::Local<v8::Value> exception = try_catch.Exception();
    lastError_ = *v8::String::Utf8Value(isolate_, exception);
}

void V8Engine::setupParameters(v8::Local<v8::Context> context, const std::map<std::string, std::string>& params) {
    v8::Local<v8::Object> paramsObj = v8::Object::New(isolate_);
    
    for (const auto& param : params) {
        v8::Local<v8::String> key = v8::String::NewFromUtf8(isolate_, param.first.c_str()).ToLocalChecked();
        v8::Local<v8::String> value = v8::String::NewFromUtf8(isolate_, param.second.c_str()).ToLocalChecked();
        paramsObj->Set(context, key, value).Check();
    }
    
    context->Global()->Set(context, 
        v8::String::NewFromUtf8(isolate_, "PARAMS").ToLocalChecked(), 
        paramsObj).Check();
}

void V8Engine::setupApiMethods(v8::Local<v8::Object> api) {
    v8::Local<v8::Context> context = isolate_->GetCurrentContext();
    
    // GetActiveSheet mock
    api->Set(context,
        v8::String::NewFromUtf8(isolate_, "GetActiveSheet").ToLocalChecked(),
        v8::Function::New(context, ApiGetActiveSheet, v8::External::New(isolate_, this)).ToLocalChecked()).Check();
    
    // GetActiveDocument mock  
    api->Set(context,
        v8::String::NewFromUtf8(isolate_, "GetActiveDocument").ToLocalChecked(),
        v8::Function::New(context, ApiGetActiveDocument, v8::External::New(isolate_, this)).ToLocalChecked()).Check();
    
    // GetActivePresentation mock
    api->Set(context,
        v8::String::NewFromUtf8(isolate_, "GetActivePresentation").ToLocalChecked(),
        v8::Function::New(context, ApiGetActivePresentation, v8::External::New(isolate_, this)).ToLocalChecked()).Check();
    
    // ShowMessage mock
    api->Set(context,
        v8::String::NewFromUtf8(isolate_, "ShowMessage").ToLocalChecked(),
        v8::Function::New(context, ApiShowMessage, v8::External::New(isolate_, this)).ToLocalChecked()).Check();
    
    // Save mock
    api->Set(context,
        v8::String::NewFromUtf8(isolate_, "Save").ToLocalChecked(),
        v8::Function::New(context, ApiSave, v8::External::New(isolate_, this)).ToLocalChecked()).Check();
}

void V8Engine::setupConsoleObject(v8::Local<v8::Context> context) {
    v8::Local<v8::Object> console = v8::Object::New(isolate_);
    
    // console.log
    console->Set(context,
        v8::String::NewFromUtf8(isolate_, "log").ToLocalChecked(),
        v8::Function::New(context, ConsoleLog, v8::External::New(isolate_, this)).ToLocalChecked()).Check();
    
    context->Global()->Set(context,
        v8::String::NewFromUtf8(isolate_, "console").ToLocalChecked(),
        console).Check();
}

void V8Engine::setupSheetMethods(v8::Local<v8::Object> sheet, v8::Isolate* isolate) {
    v8::Local<v8::Context> context = isolate->GetCurrentContext();
    
    // GetRange mock
    sheet->Set(context,
        v8::String::NewFromUtf8(isolate, "GetRange").ToLocalChecked(),
        v8::Function::New(context, SheetGetRange, v8::External::New(isolate, nullptr)).ToLocalChecked()).Check();
}

void V8Engine::setupRangeMethods(v8::Local<v8::Object> range, v8::Isolate* isolate) {
    v8::Local<v8::Context> context = isolate->GetCurrentContext();
    
    // SetValue mock
    range->Set(context,
        v8::String::NewFromUtf8(isolate, "SetValue").ToLocalChecked(),
        v8::Function::New(context, RangeSetValue, v8::External::New(isolate, nullptr)).ToLocalChecked()).Check();
    
    // GetValue mock
    range->Set(context,
        v8::String::NewFromUtf8(isolate, "GetValue").ToLocalChecked(),
        v8::Function::New(context, RangeGetValue, v8::External::New(isolate, nullptr)).ToLocalChecked()).Check();
}

void V8Engine::setupDocumentMethods(v8::Local<v8::Object> document, v8::Isolate* isolate) {
    v8::Local<v8::Context> context = isolate->GetCurrentContext();
    
    // CreateParagraph mock
    document->Set(context,
        v8::String::NewFromUtf8(isolate, "CreateParagraph").ToLocalChecked(),
        v8::Function::New(context, [](const v8::FunctionCallbackInfo<v8::Value>& args) {
            v8::Local<v8::Object> paragraph = v8::Object::New(args.GetIsolate());
            args.GetReturnValue().Set(paragraph);
        }).ToLocalChecked()).Check();
}

void V8Engine::setupPresentationMethods(v8::Local<v8::Object> presentation, v8::Isolate* isolate) {
    v8::Local<v8::Context> context = isolate->GetCurrentContext();
    
    // GetSlideByIndex mock
    presentation->Set(context,
        v8::String::NewFromUtf8(isolate, "GetSlideByIndex").ToLocalChecked(),
        v8::Function::New(context, [](const v8::FunctionCallbackInfo<v8::Value>& args) {
            v8::Local<v8::Object> slide = v8::Object::New(args.GetIsolate());
            args.GetReturnValue().Set(slide);
        }).ToLocalChecked()).Check();
}

// Static callback implementations
void V8Engine::ApiGetActiveSheet(const v8::FunctionCallbackInfo<v8::Value>& args) {
    v8::Local<v8::Object> sheet = v8::Object::New(args.GetIsolate());
    setupSheetMethods(sheet, args.GetIsolate());
    args.GetReturnValue().Set(sheet);
}

void V8Engine::ApiGetActiveDocument(const v8::FunctionCallbackInfo<v8::Value>& args) {
    v8::Local<v8::Object> document = v8::Object::New(args.GetIsolate());
    setupDocumentMethods(document, args.GetIsolate());
    args.GetReturnValue().Set(document);
}

void V8Engine::ApiGetActivePresentation(const v8::FunctionCallbackInfo<v8::Value>& args) {
    v8::Local<v8::Object> presentation = v8::Object::New(args.GetIsolate());
    setupPresentationMethods(presentation, args.GetIsolate());
    args.GetReturnValue().Set(presentation);
}

void V8Engine::ApiShowMessage(const v8::FunctionCallbackInfo<v8::Value>& args) {
    if (args.Length() >= 2) {
        v8::String::Utf8Value title(args.GetIsolate(), args[0]);
        v8::String::Utf8Value message(args.GetIsolate(), args[1]);
        printf("📋 Api.ShowMessage: %s - %s\n", *title, *message);
    }
}

void V8Engine::ApiSave(const v8::FunctionCallbackInfo<v8::Value>& args) {
    printf("💾 Api.Save: Document saved\n");
}

void V8Engine::ConsoleLog(const v8::FunctionCallbackInfo<v8::Value>& args) {
    for (int i = 0; i < args.Length(); i++) {
        v8::String::Utf8Value str(args.GetIsolate(), args[i]);
        printf("%s%s", i > 0 ? " " : "", *str);
    }
    printf("\n");
}

void V8Engine::RangeSetValue(const v8::FunctionCallbackInfo<v8::Value>& args) {
    if (args.Length() >= 1) {
        v8::String::Utf8Value value(args.GetIsolate(), args[0]);
        printf("📝 Range.SetValue: %s\n", *value);
    }
}

void V8Engine::RangeGetValue(const v8::FunctionCallbackInfo<v8::Value>& args) {
    printf("📖 Range.GetValue: MockValue\n");
    args.GetReturnValue().Set(v8::String::NewFromUtf8(args.GetIsolate(), "MockValue").ToLocalChecked());
}

void V8Engine::SheetGetRange(const v8::FunctionCallbackInfo<v8::Value>& args) {
    v8::Local<v8::Object> range = v8::Object::New(args.GetIsolate());
    setupRangeMethods(range, args.GetIsolate());
    
    if (args.Length() >= 1) {
        v8::String::Utf8Value rangeAddr(args.GetIsolate(), args[0]);
        printf("📊 Sheet.GetRange: %s\n", *rangeAddr);
    }
    
    args.GetReturnValue().Set(range);
}

} // namespace macro
} // namespace onlyoffice

#endif // HAVE_V8