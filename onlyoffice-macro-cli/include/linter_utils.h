#pragma once

#include "linter.h"
#include <string>
#include <vector>

namespace onlyoffice {
namespace macro {

std::string getLineContent(const std::string& source, int lineNumber);
std::vector<std::string> splitLines(const std::string& source);
void addIssue(std::vector<LintIssue>& issues, int line, int column, 
              const std::string& message, const std::string& rule, 
              LintSeverity severity, const std::string& source);

} // namespace macro
} // namespace onlyoffice