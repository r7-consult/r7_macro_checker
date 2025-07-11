#include "linter_rules.h"
#include "linter_utils.h"
#include <regex>
#include <set>

namespace onlyoffice {
namespace macro {

void checkFunctionCalls(const std::string& source, std::vector<LintIssue>& issues,
                       const std::set<std::string>& knownGlobals) {
    auto lines = splitLines(source);
    
    for (size_t i = 0; i < lines.size(); i++) {
        const std::string& line = lines[i];
        int lineNum = static_cast<int>(i + 1);
        
        // Check for malformed function calls
        // Updated regex to handle chained method calls properly
        std::regex funcCallRegex(R"(([a-zA-Z_$][a-zA-Z0-9_$]*(?:\.[a-zA-Z_$][a-zA-Z0-9_$]*)*)\s*\()");
        std::sregex_iterator iter(line.begin(), line.end(), funcCallRegex);
        std::sregex_iterator end;
        
        for (; iter != end; ++iter) {
            const std::smatch& match = *iter;
            std::string funcName = match[1].str();
            
            // Check for common typos in function names
            if (funcName.find("console.log") == 0 && funcName != "console.log") {
                addIssue(issues, lineNum, static_cast<int>(match.position()), 
                        "Invalid function call '" + funcName + "', did you mean 'console.log'?",
                        "functions", LintSeverity::Error, source);
            }
            
            // Improved context awareness for function vs language construct detection
            bool isMethodCall = funcName.find('.') != std::string::npos;
            
            // Check if this is a chained method call (e.g., obj.method().anotherMethod())
            bool isChainedMethodCall = false;
            if (!isMethodCall) {
                // Look for pattern like ").functionName(" or ".functionName(" indicating a chained method call
                size_t pos = match.position();
                if (pos > 0 && (line[pos - 1] == ')' || line[pos - 1] == '.')) {
                    isChainedMethodCall = true;
                }
            }
            
            // Skip checking for method calls and chained method calls (they're usually safe)
            if (isMethodCall || isChainedMethodCall) {
                continue;
            }
            
            // Check for undefined functions (only for standalone function calls)
            if (knownGlobals.find(funcName) == knownGlobals.end() &&
                funcName != "function" && funcName != "return") {
                
                // Check if it's defined in the source
                std::regex funcDefRegex("function\\s+" + funcName + "\\s*\\(");
                if (!std::regex_search(source, funcDefRegex)) {
                    // Additional check: skip if it's a JavaScript language construct
                    // (This handles cases where the regex misidentifies keywords)
                    if (funcName != "if" && funcName != "else" && funcName != "for" && 
                        funcName != "while" && funcName != "do" && funcName != "switch" &&
                        funcName != "try" && funcName != "catch" && funcName != "finally" &&
                        funcName != "throw" && funcName != "typeof" && funcName != "new" &&
                        funcName != "delete" && funcName != "in" && funcName != "instanceof" &&
                        funcName != "var" && funcName != "let" && funcName != "const" &&
                        funcName != "break" && funcName != "continue") {
                        
                        addIssue(issues, lineNum, static_cast<int>(match.position()), 
                                "Function '" + funcName + "' may not be defined",
                                "functions", LintSeverity::Warning, source);
                    }
                }
            }
        }
    }
}

} // namespace macro
} // namespace onlyoffice