#pragma once

#include "linter.h"
#include <string>
#include <vector>
#include <set>

namespace onlyoffice {
namespace macro {

void checkBasicSyntax(const std::string& source, std::vector<LintIssue>& issues);
void checkVariableDeclarations(const std::string& source, std::vector<LintIssue>& issues, 
                               const std::set<std::string>& knownGlobals);
void checkFunctionCalls(const std::string& source, std::vector<LintIssue>& issues,
                       const std::set<std::string>& knownGlobals);
void checkOnlyOfficeAPI(const std::string& source, std::vector<LintIssue>& issues,
                       const std::set<std::string>& knownAPIs);
void checkCommonMistakes(const std::string& source, std::vector<LintIssue>& issues);
void checkCodeStyle(const std::string& source, std::vector<LintIssue>& issues);

} // namespace macro
} // namespace onlyoffice