#include "linter_rules.h"
#include "linter_utils.h"

namespace onlyoffice {
namespace macro {

void checkCodeStyle(const std::string& source, std::vector<LintIssue>& issues) {
    auto lines = splitLines(source);
    
    for (size_t i = 0; i < lines.size(); i++) {
        const std::string& line = lines[i];
        int lineNum = static_cast<int>(i + 1);
        
        // Check line length
        if (line.length() > 120) {
            addIssue(issues, lineNum, 120, 
                    "Line too long (" + std::to_string(line.length()) + " characters)",
                    "code-style", LintSeverity::Info, source);
        }
        
        // Check for trailing whitespace
        if (!line.empty() && (line.back() == ' ' || line.back() == '\t')) {
            addIssue(issues, lineNum, static_cast<int>(line.length() - 1), 
                    "Trailing whitespace",
                    "code-style", LintSeverity::Info, source);
        }
        
        // Check for mixed tabs and spaces (basic check)
        if (line.find('\t') != std::string::npos && line.find(' ') != std::string::npos) {
            addIssue(issues, lineNum, 0, 
                    "Mixed tabs and spaces for indentation",
                    "code-style", LintSeverity::Info, source);
        }
    }
}

} // namespace macro
} // namespace onlyoffice