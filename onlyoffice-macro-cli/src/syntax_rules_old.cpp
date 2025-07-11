#include "linter_rules.h"
#include "linter_utils.h"

extern "C" {
#include "duktape.h"
}

namespace onlyoffice {
namespace macro {

void checkBasicSyntax(const std::string& source, std::vector<LintIssue>& issues) {
    // Use Duktape to check basic syntax
    duk_context* ctx = duk_create_heap_default();
    if (!ctx) {
        addIssue(issues, 0, 0, "Failed to create JavaScript context for syntax checking", 
                "syntax", LintSeverity::Error, source);
        return;
    }
    
    duk_push_string(ctx, source.c_str());
    duk_push_string(ctx, "lint-check");
    
    if (duk_pcompile(ctx, 0) != 0) {
        if (duk_is_error(ctx, -1)) {
            duk_get_prop_string(ctx, -1, "lineNumber");
            int line = duk_get_int_default(ctx, -1, 0);
            duk_pop(ctx);
            
            std::string message = duk_safe_to_string(ctx, -1);
            addIssue(issues, line, 0, message, "syntax", LintSeverity::Error, source);
        } else {
            std::string message = duk_safe_to_string(ctx, -1);
            addIssue(issues, 0, 0, message, "syntax", LintSeverity::Error, source);
        }
    }
    
    duk_destroy_heap(ctx);
}

} // namespace macro
} // namespace onlyoffice