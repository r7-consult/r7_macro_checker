#include "linter_utils.h"
#include <sstream>

namespace onlyoffice {
namespace macro {

std::string getLineContent(const std::string& source, int lineNumber) {
    auto lines = splitLines(source);
    if (lineNumber > 0 && lineNumber <= static_cast<int>(lines.size())) {
        return lines[lineNumber - 1];
    }
    return "";
}

std::vector<std::string> splitLines(const std::string& source) {
    std::vector<std::string> lines;
    std::istringstream stream(source);
    std::string line;
    
    while (std::getline(stream, line)) {
        lines.push_back(line);
    }
    
    return lines;
}

void addIssue(std::vector<LintIssue>& issues, int line, int column, 
              const std::string& message, const std::string& rule, 
              LintSeverity severity, const std::string& source) {
    LintIssue issue;
    issue.line = line;
    issue.column = column;
    issue.message = message;
    issue.rule = rule;
    issue.severity = severity;
    issue.codeContext = getLineContent(source, line);
    
    issues.push_back(issue);
}

} // namespace macro
} // namespace onlyoffice