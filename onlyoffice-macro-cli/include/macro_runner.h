#pragma once

#include <string>
#include <memory>
#include <map>
#include "js_engine_interface.h"
#include "syntax_checker.h"
#include "linter.h"

namespace onlyoffice {
namespace macro {

enum class ExecutionResult {
    Success,
    SyntaxError,
    RuntimeError,
    FileNotFound,
    InvalidArguments
};

struct MacroExecutionInfo {
    ExecutionResult result;
    std::string message;
    double executionTime;
    SyntaxCheckResult syntaxResult;
    LintResult lintResult;
};

class MacroRunner {
public:
    MacroRunner();
    MacroRunner(std::unique_ptr<JSEngineInterface> engine);
    ~MacroRunner();
    
    MacroExecutionInfo runMacroFile(const std::string& filepath);
    MacroExecutionInfo runMacroString(const std::string& source);
    
    void setVerbose(bool verbose);
    void setSyntaxCheckOnly(bool syntaxOnly);
    void setDryRun(bool dryRun);
    void setDocumentPath(const std::string& path);
    void setLintingEnabled(bool enabled);
    void setStrictLinting(bool strict);
    
    void addParameter(const std::string& key, const std::string& value);
    void clearParameters();
    
    SyntaxCheckResult checkSyntax(const std::string& source);
    
private:
    std::unique_ptr<JSEngineInterface> engine;
    std::unique_ptr<SyntaxChecker> syntaxChecker;
    std::unique_ptr<Linter> linter;
    
    bool verbose;
    bool syntaxOnly;
    bool dryRun;
    bool lintingEnabled;
    bool strictLinting;
    std::string documentPath;
    std::map<std::string, std::string> parameters;
    
    void logMessage(const std::string& message);
    void logError(const std::string& error);
    void setupEngine();
    void setupParameters();
    double getCurrentTime();
};

} // namespace macro
} // namespace onlyoffice