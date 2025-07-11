#include "macro_runner.h"
#include <fstream>
#include <sstream>
#include <chrono>
#include <iostream>
#include <map>

namespace onlyoffice {
namespace macro {

MacroRunner::MacroRunner() 
    : engine(nullptr)
    , syntaxChecker(std::make_unique<SyntaxChecker>())
    , linter(std::make_unique<Linter>())
    , verbose(false)
    , syntaxOnly(false)
    , dryRun(false)
    , lintingEnabled(true)
    , strictLinting(false) {
    
    // Create default engine if none provided
    engine = JSEngineFactory::createEngine(JSEngineFactory::EngineType::Auto);
    if (engine) {
        engine->initialize();
        engine->setupOnlyOfficeAPI();
    }
}

MacroRunner::MacroRunner(std::unique_ptr<JSEngineInterface> jsEngine)
    : engine(std::move(jsEngine))
    , syntaxChecker(std::make_unique<SyntaxChecker>())
    , linter(std::make_unique<Linter>())
    , verbose(false)
    , syntaxOnly(false)
    , dryRun(false)
    , lintingEnabled(true)
    , strictLinting(false) {
}

MacroRunner::~MacroRunner() = default;

MacroExecutionInfo MacroRunner::runMacroFile(const std::string& filepath) {
    MacroExecutionInfo info;
    auto start = getCurrentTime();
    
    // Check if file exists
    std::ifstream file(filepath);
    if (!file.is_open()) {
        info.result = ExecutionResult::FileNotFound;
        info.message = "Cannot open file: " + filepath;
        info.executionTime = getCurrentTime() - start;
        return info;
    }
    
    // Read file content
    std::ostringstream buffer;
    buffer << file.rdbuf();
    std::string source = buffer.str();
    file.close();
    
    logMessage("Running macro file: " + filepath);
    
    return runMacroString(source);
}

MacroExecutionInfo MacroRunner::runMacroString(const std::string& source) {
    MacroExecutionInfo info;
    auto start = getCurrentTime();
    
    // Run linter first if enabled
    if (lintingEnabled) {
        info.lintResult = linter->lintString(source);
        
        if (info.lintResult.hasErrors()) {
            info.result = ExecutionResult::SyntaxError;
            info.message = "Linting errors found";
            info.executionTime = getCurrentTime() - start;
            return info;
        }
    }
    
    // Use engine for syntax validation instead of old syntax checker
    std::vector<LintIssue> syntaxIssues;
    if (engine && !engine->validateSyntax(source, syntaxIssues)) {
        info.result = ExecutionResult::SyntaxError;
        info.message = "Syntax errors found";
        info.executionTime = getCurrentTime() - start;
        
        // Convert engine syntax issues to old format for compatibility
        for (const auto& issue : syntaxIssues) {
            SyntaxError syntaxError;
            syntaxError.line = issue.line;
            syntaxError.column = issue.column;
            syntaxError.message = issue.message;
            syntaxError.severity = (issue.severity == LintSeverity::Error) ? "error" : "warning";
            info.syntaxResult.errors.push_back(syntaxError);
        }
        info.syntaxResult.isValid = false;
        return info;
    }
    
    info.syntaxResult.isValid = true;
    
    // If syntax-only mode, return success
    if (syntaxOnly) {
        info.result = ExecutionResult::Success;
        info.message = "Syntax check passed";
        info.executionTime = getCurrentTime() - start;
        return info;
    }
    
    // If dry-run mode, return success without execution
    if (dryRun) {
        info.result = ExecutionResult::Success;
        info.message = "Dry run completed successfully";
        info.executionTime = getCurrentTime() - start;
        logMessage("Dry run completed - syntax is valid");
        return info;
    }
    
    // Setup parameters before execution
    setupParameters();
    
    // Execute the macro
    logMessage("Executing macro...");
    
    if (!engine->executeScript(source, parameters)) {
        info.result = ExecutionResult::RuntimeError;
        info.message = engine->getLastError();
        info.executionTime = getCurrentTime() - start;
        return info;
    }
    
    info.result = ExecutionResult::Success;
    info.message = "Macro executed successfully";
    info.executionTime = getCurrentTime() - start;
    
    logMessage("Macro execution completed");
    
    return info;
}

void MacroRunner::setVerbose(bool verbose) {
    this->verbose = verbose;
}

void MacroRunner::setSyntaxCheckOnly(bool syntaxOnly) {
    this->syntaxOnly = syntaxOnly;
}

void MacroRunner::setDryRun(bool dryRun) {
    this->dryRun = dryRun;
}

void MacroRunner::setDocumentPath(const std::string& path) {
    this->documentPath = path;
    
    // Store document path for use in parameters
    parameters["__DOCUMENT_PATH__"] = path;
}

void MacroRunner::setLintingEnabled(bool enabled) {
    this->lintingEnabled = enabled;
}

void MacroRunner::setStrictLinting(bool strict) {
    this->strictLinting = strict;
    if (linter) {
        linter->setStrictMode(strict);
    }
}

void MacroRunner::addParameter(const std::string& key, const std::string& value) {
    parameters[key] = value;
    
    // Set parameter in engine
    if (engine) {
        engine->setGlobalString("__PARAM_" + key + "__", value);
    }
}

void MacroRunner::clearParameters() {
    parameters.clear();
}

SyntaxCheckResult MacroRunner::checkSyntax(const std::string& source) {
    SyntaxCheckResult result;
    result.source = source;
    
    if (engine) {
        std::vector<LintIssue> issues;
        result.isValid = engine->validateSyntax(source, issues);
        
        // Convert engine issues to syntax errors
        for (const auto& issue : issues) {
            SyntaxError syntaxError;
            syntaxError.line = issue.line;
            syntaxError.column = issue.column;
            syntaxError.message = issue.message;
            syntaxError.severity = (issue.severity == LintSeverity::Error) ? "error" : "warning";
            result.errors.push_back(syntaxError);
        }
    } else {
        // Fallback to old syntax checker if no engine available
        return syntaxChecker->checkString(source);
    }
    
    return result;
}

void MacroRunner::logMessage(const std::string& message) {
    if (verbose) {
        std::cout << "[INFO] " << message << std::endl;
    }
}

void MacroRunner::logError(const std::string& error) {
    if (verbose) {
        std::cerr << "[ERROR] " << error << std::endl;
    }
}

void MacroRunner::setupParameters() {
    // Parameters are now passed directly to executeScript method
    // No need for separate setup
    logMessage("Parameters ready for execution");
}

double MacroRunner::getCurrentTime() {
    auto now = std::chrono::high_resolution_clock::now();
    auto duration = now.time_since_epoch();
    return std::chrono::duration_cast<std::chrono::milliseconds>(duration).count();
}

} // namespace macro
} // namespace onlyoffice