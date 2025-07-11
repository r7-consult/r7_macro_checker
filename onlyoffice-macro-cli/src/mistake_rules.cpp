#include "linter_rules.h"
#include "linter_utils.h"
#include <regex>

namespace onlyoffice {
namespace macro {

void checkCommonMistakes(const std::string& source, std::vector<LintIssue>& issues) {
    auto lines = splitLines(source);
    
    for (size_t i = 0; i < lines.size(); i++) {
        const std::string& line = lines[i];
        int lineNum = static_cast<int>(i + 1);
        
        // Check for assignment in conditions
        std::regex assignInCondRegex(R"(if\s*\([^)]*=\s*[^=])");
        if (std::regex_search(line, assignInCondRegex)) {
            addIssue(issues, lineNum, static_cast<int>(line.find("if")), 
                    "Assignment in condition, did you mean '==' or '==='?",
                    "common-mistakes", LintSeverity::Warning, source);
        }
        
        // Check for missing semicolons (basic check)
        if (!line.empty() && 
            line.find("//") == std::string::npos && 
            line.find("/*") == std::string::npos &&
            line.back() != ';' && 
            line.back() != '}' && 
            line.back() != '{' &&
            line.find("function") == std::string::npos &&
            line.find("if") == std::string::npos &&
            line.find("for") == std::string::npos &&
            line.find("while") == std::string::npos) {
            
            std::regex stmtRegex(R"(\w+\s*[\(\)\[\]=+\-*/]\s*\w*)");
            if (std::regex_search(line, stmtRegex)) {
                addIssue(issues, lineNum, static_cast<int>(line.length()), 
                        "Missing semicolon",
                        "common-mistakes", LintSeverity::Info, source);
            }
        }
        
        // Check for == vs ===
        if (line.find("==") != std::string::npos && line.find("===") == std::string::npos) {
            addIssue(issues, lineNum, static_cast<int>(line.find("==")), 
                    "Consider using '===' instead of '==' for strict equality",
                    "common-mistakes", LintSeverity::Info, source);
        }
    }
}

} // namespace macro
} // namespace onlyoffice