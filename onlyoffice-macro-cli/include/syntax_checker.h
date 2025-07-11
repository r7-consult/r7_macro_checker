#pragma once

#include <string>
#include <vector>
#include <memory>

namespace onlyoffice {
namespace macro {

struct SyntaxError {
    int line;
    int column;
    std::string message;
    std::string severity; // "error", "warning", "info"
};

struct SyntaxCheckResult {
    bool isValid;
    std::vector<SyntaxError> errors;
    std::string source;
    
    bool hasErrors() const {
        for (const auto& error : errors) {
            if (error.severity == "error") {
                return true;
            }
        }
        return false;
    }
    
    bool hasWarnings() const {
        for (const auto& error : errors) {
            if (error.severity == "warning") {
                return true;
            }
        }
        return false;
    }
};

class SyntaxChecker {
public:
    SyntaxChecker();
    ~SyntaxChecker();
    
    SyntaxCheckResult checkFile(const std::string& filepath);
    SyntaxCheckResult checkString(const std::string& source);
    
    void setStrictMode(bool strict);
    void setOnlyOfficeAPIChecks(bool enable);
    
private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
    
    bool validateOnlyOfficeAPI(const std::string& source, std::vector<SyntaxError>& errors);
    bool checkJavaScriptSyntax(const std::string& source, std::vector<SyntaxError>& errors);
    void addKnownAPIs();
};

} // namespace macro
} // namespace onlyoffice