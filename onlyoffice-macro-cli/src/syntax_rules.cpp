#include "linter_rules.h"
#include "linter_utils.h"
#include "js_engine_interface.h"

namespace onlyoffice {
namespace macro {

void checkBasicSyntax(const std::string& source, std::vector<LintIssue>& issues) {
    // Use the best available JavaScript engine for syntax checking
    auto engine = JSEngineFactory::createEngine(JSEngineFactory::EngineType::Auto);
    
    if (!engine) {
        addIssue(issues, 0, 0, "No JavaScript engine available for syntax checking", 
                "syntax", LintSeverity::Error, source);
        return;
    }
    
    if (!engine->initialize()) {
        addIssue(issues, 0, 0, "Failed to initialize JavaScript engine: " + engine->getLastError(), 
                "syntax", LintSeverity::Error, source);
        return;
    }
    
    // Use engine to validate syntax
    std::vector<LintIssue> syntaxIssues;
    if (!engine->validateSyntax(source, syntaxIssues)) {
        // Add all syntax issues to the main issues list
        issues.insert(issues.end(), syntaxIssues.begin(), syntaxIssues.end());
    }
    
    engine->cleanup();
}

} // namespace macro
} // namespace onlyoffice