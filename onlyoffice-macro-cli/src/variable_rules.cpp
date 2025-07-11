#include "linter_rules.h"
#include "linter_utils.h"
#include <regex>
#include <set>

namespace onlyoffice {
namespace macro {

void checkVariableDeclarations(const std::string& source, std::vector<LintIssue>& issues, 
                               const std::set<std::string>& knownGlobals) {
    auto lines = splitLines(source);
    
    for (size_t i = 0; i < lines.size(); i++) {
        const std::string& line = lines[i];
        int lineNum = static_cast<int>(i + 1);
        
        // Check for variables declared without var/let/const
        std::regex undeclaredVarRegex(R"(^\s*([a-zA-Z_$][a-zA-Z0-9_$]*)\s*=\s*[^=])");
        std::smatch match;
        if (std::regex_search(line, match, undeclaredVarRegex)) {
            // Skip if it's inside a comment or string
            size_t commentPos1 = line.find("//");
            size_t commentPos2 = line.find("/*");
            if ((commentPos1 != std::string::npos && commentPos1 < static_cast<size_t>(match.position())) || 
                (commentPos2 != std::string::npos && commentPos2 < static_cast<size_t>(match.position()))) {
                continue;
            }
            
            std::string varName = match[1].str();
            if (knownGlobals.find(varName) == knownGlobals.end()) {
                addIssue(issues, lineNum, static_cast<int>(match.position()), 
                        "Variable '" + varName + "' should be declared with 'var', 'let', or 'const'",
                        "variables", LintSeverity::Warning, source);
            }
        }
        
        // Check for unused variables (basic check)
        std::regex varDeclRegex(R"((var|let|const)\s+([a-zA-Z_$][a-zA-Z0-9_$]*))");
        std::sregex_iterator iter(line.begin(), line.end(), varDeclRegex);
        std::sregex_iterator end;
        
        for (; iter != end; ++iter) {
            const std::smatch& varMatch = *iter;
            std::string varName = varMatch[2].str();
            
            // Simple check: if variable is not used in the rest of the source
            if (source.find(varName, match.position() + varName.length()) == std::string::npos) {
                addIssue(issues, lineNum, static_cast<int>(varMatch.position()), 
                        "Variable '" + varName + "' is declared but never used",
                        "variables", LintSeverity::Warning, source);
            }
        }
    }
}

} // namespace macro
} // namespace onlyoffice