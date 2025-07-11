#pragma once

#include <string>
#include <vector>
#include <memory>
#include <regex>

namespace onlyoffice {
namespace macro {

enum class LintSeverity {
    Error,
    Warning,
    Info
};

struct LintIssue {
    int line;
    int column;
    std::string message;
    std::string rule;
    LintSeverity severity;
    std::string codeContext;
};

struct LintResult {
    bool isValid;
    std::vector<LintIssue> issues;
    
    bool hasErrors() const {
        for (const auto& issue : issues) {
            if (issue.severity == LintSeverity::Error) {
                return true;
            }
        }
        return false;
    }
    
    bool hasWarnings() const {
        for (const auto& issue : issues) {
            if (issue.severity == LintSeverity::Warning) {
                return true;
            }
        }
        return false;
    }
    
    int getErrorCount() const {
        int count = 0;
        for (const auto& issue : issues) {
            if (issue.severity == LintSeverity::Error) count++;
        }
        return count;
    }
    
    int getWarningCount() const {
        int count = 0;
        for (const auto& issue : issues) {
            if (issue.severity == LintSeverity::Warning) count++;
        }
        return count;
    }
};

class Linter {
public:
    Linter();
    ~Linter();
    
    LintResult lintFile(const std::string& filepath);
    LintResult lintString(const std::string& source);
    
    // Configuration
    void setStrictMode(bool strict);
    void setOnlyOfficeMode(bool enable);
    void enableRule(const std::string& rule);
    void disableRule(const std::string& rule);
    
private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace macro
} // namespace onlyoffice