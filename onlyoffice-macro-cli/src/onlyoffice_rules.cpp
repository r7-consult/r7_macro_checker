#include "linter_rules.h"
#include "linter_utils.h"
#include <regex>
#include <set>

namespace onlyoffice {
namespace macro {

void checkOnlyOfficeAPI(const std::string& source, std::vector<LintIssue>& issues,
                       const std::set<std::string>& knownAPIs) {
    auto lines = splitLines(source);
    
    for (size_t i = 0; i < lines.size(); i++) {
        const std::string& line = lines[i];
        int lineNum = static_cast<int>(i + 1);
        
        // Check for OnlyOffice API usage
        std::regex apiCallRegex(R"((Api(?:\.[a-zA-Z_$][a-zA-Z0-9_$]*)*)\s*\()");
        std::sregex_iterator iter(line.begin(), line.end(), apiCallRegex);
        std::sregex_iterator end;
        
        for (; iter != end; ++iter) {
            const std::smatch& match = *iter;
            std::string apiCall = match[1].str();
            
            if (knownAPIs.find(apiCall) == knownAPIs.end()) {
                addIssue(issues, lineNum, static_cast<int>(match.position()), 
                        "Unknown OnlyOffice API call: " + apiCall,
                        "onlyoffice-api", LintSeverity::Warning, source);
            }
        }
        
        // Check for browser-specific APIs that don't work in OnlyOffice
        if (line.find("document.") != std::string::npos) {
            addIssue(issues, lineNum, static_cast<int>(line.find("document.")), 
                    "Use 'Api.GetActiveDocument()' instead of 'document' in OnlyOffice macros",
                    "onlyoffice-api", LintSeverity::Warning, source);
        }
        
        if (line.find("window.") != std::string::npos) {
            addIssue(issues, lineNum, static_cast<int>(line.find("window.")), 
                    "'window' object is not available in OnlyOffice macros",
                    "onlyoffice-api", LintSeverity::Warning, source);
        }
        
        if (line.find("alert(") != std::string::npos) {
            addIssue(issues, lineNum, static_cast<int>(line.find("alert(")), 
                    "Use 'Api.ShowMessage()' instead of 'alert()' in OnlyOffice macros",
                    "onlyoffice-api", LintSeverity::Warning, source);
        }
    }
}

} // namespace macro
} // namespace onlyoffice