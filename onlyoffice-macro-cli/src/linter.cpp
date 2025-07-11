#include "linter.h"
#include "linter_rules.h"
#include "linter_utils.h"
#include "api_definitions.h"
#include <fstream>
#include <sstream>
#include <set>
#include <algorithm>

namespace onlyoffice {
namespace macro {

class Linter::Impl {
public:
    bool strictMode = false;
    bool onlyOfficeMode = true;
    std::set<std::string> enabledRules;
    std::set<std::string> disabledRules;
    
    // Known OnlyOffice API methods
    std::set<std::string> knownAPIs;
    
    // Known global functions that should exist
    std::set<std::string> knownGlobals;
    
    Impl() {
        setupKnownAPIs();
        setupDefaultRules();
    }
    
    void setupKnownAPIs() {
        onlyoffice::macro::setupKnownAPIs(knownAPIs);
        onlyoffice::macro::setupKnownGlobals(knownGlobals);
    }
    
    void setupDefaultRules() {
        enabledRules.insert("syntax");
        enabledRules.insert("variables");
        enabledRules.insert("functions");
        enabledRules.insert("onlyoffice-api");
        enabledRules.insert("common-mistakes");
        enabledRules.insert("code-style");
    }
    
    bool isRuleEnabled(const std::string& rule) {
        if (disabledRules.count(rule)) return false;
        return enabledRules.count(rule) > 0;
    }
};

Linter::Linter() : pImpl(std::make_unique<Impl>()) {}

Linter::~Linter() = default;

LintResult Linter::lintFile(const std::string& filepath) {
    std::ifstream file(filepath);
    if (!file.is_open()) {
        LintResult result;
        result.isValid = false;
        LintIssue issue;
        issue.line = 0;
        issue.column = 0;
        issue.message = "Cannot open file: " + filepath;
        issue.rule = "file-access";
        issue.severity = LintSeverity::Error;
        result.issues.push_back(issue);
        return result;
    }
    
    std::ostringstream buffer;
    buffer << file.rdbuf();
    std::string source = buffer.str();
    
    return lintString(source);
}

LintResult Linter::lintString(const std::string& source) {
    LintResult result;
    result.isValid = true;
    
    // Run all enabled lint rules
    if (pImpl->isRuleEnabled("syntax")) {
        checkBasicSyntax(source, result.issues);
    }
    
    if (pImpl->isRuleEnabled("variables")) {
        checkVariableDeclarations(source, result.issues, pImpl->knownGlobals);
    }
    
    if (pImpl->isRuleEnabled("functions")) {
        checkFunctionCalls(source, result.issues, pImpl->knownGlobals);
    }
    
    if (pImpl->isRuleEnabled("onlyoffice-api") && pImpl->onlyOfficeMode) {
        checkOnlyOfficeAPI(source, result.issues, pImpl->knownAPIs);
    }
    
    if (pImpl->isRuleEnabled("common-mistakes")) {
        checkCommonMistakes(source, result.issues);
    }
    
    if (pImpl->isRuleEnabled("code-style")) {
        checkCodeStyle(source, result.issues);
    }
    
    // Check if there are any errors
    if (result.hasErrors()) {
        result.isValid = false;
    }
    
    return result;
}

void Linter::setStrictMode(bool strict) {
    pImpl->strictMode = strict;
}

void Linter::setOnlyOfficeMode(bool enable) {
    pImpl->onlyOfficeMode = enable;
}

void Linter::enableRule(const std::string& rule) {
    pImpl->enabledRules.insert(rule);
    pImpl->disabledRules.erase(rule);
}

void Linter::disableRule(const std::string& rule) {
    pImpl->disabledRules.insert(rule);
    pImpl->enabledRules.erase(rule);
}

} // namespace macro
} // namespace onlyoffice