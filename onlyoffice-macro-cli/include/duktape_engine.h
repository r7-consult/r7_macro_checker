#pragma once

#include <string>
#include <memory>
#include <functional>

extern "C" {
#include "duktape.h"
}

namespace onlyoffice {
namespace macro {

class DuktapeEngine {
public:
    DuktapeEngine();
    ~DuktapeEngine();
    
    bool initialize();
    void cleanup();
    
    bool executeScript(const std::string& script);
    bool executeFile(const std::string& filepath);
    bool executeScriptWithContext(const std::string& script, const std::string& source);
    
    void setGlobalObject(const std::string& name, duk_c_function func);
    void setGlobalString(const std::string& name, const std::string& value);
    void setGlobalNumber(const std::string& name, double value);
    void setGlobalBoolean(const std::string& name, bool value);
    
    std::string getLastError() const;
    bool hasError() const;
    
    // OnlyOffice API Mock functions
    void setupOnlyOfficeAPI();
    
private:
    duk_context* ctx;
    std::string lastError;
    std::string currentSource;
    bool initialized;
    
    static duk_ret_t js_print(duk_context* ctx);
    static duk_ret_t js_api_mock(duk_context* ctx);
    static duk_ret_t js_get_active_sheet(duk_context* ctx);
    static duk_ret_t js_get_range(duk_context* ctx);
    static duk_ret_t js_set_fill_color(duk_context* ctx);
    static duk_ret_t js_show_message(duk_context* ctx);
    
    void pushOnlyOfficeAPI();
    bool loadScript(const std::string& script);
    std::string readFile(const std::string& filepath);
    std::string getErrorContext(const std::string& source, int lineNumber);
};

} // namespace macro
} // namespace onlyoffice