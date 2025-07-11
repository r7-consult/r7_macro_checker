#include <iostream>
#include <chrono>
#include <iomanip>

#include "cli_parser.h"
#include "macro_runner.h"
#include "js_engine_interface.h"
#include "linter.h"

using namespace onlyoffice::macro;

int main(int argc, char* argv[]) {
    try {
        CLIParser parser;
        CLIOptions options = parser.parse(argc, argv);
        
        if (options.help) {
            parser.printHelp();
            return 0;
        }
        
        if (options.version) {
            parser.printVersion();
            return 0;
        }
        
        if (options.inputFile.empty()) {
            std::cerr << "Error: Input file is required" << std::endl;
            parser.printHelp();
            return 1;
        }
        
        // Create JavaScript engine
        auto engine = JSEngineFactory::createEngine(JSEngineFactory::EngineType::Auto);
        if (!engine) {
            std::cerr << "Error: No JavaScript engine available" << std::endl;
            return 99;
        }
        
        if (!engine->initialize()) {
            std::cerr << "Error: Failed to initialize JavaScript engine: " << engine->getLastError() << std::endl;
            return 99;
        }
        
        // Setup OnlyOffice API
        engine->setupOnlyOfficeAPI();
        
        if (options.verbose) {
            std::cout << "Using " << engine->getEngineName() << " " << engine->getEngineVersion() << " engine" << std::endl;
        }
        
        // Create macro runner with new engine
        MacroRunner runner(std::move(engine));
        runner.setVerbose(options.verbose);
        runner.setSyntaxCheckOnly(options.syntaxCheck || options.lintOnly);
        runner.setDryRun(options.dryRun);
        runner.setLintingEnabled(!options.disableLinting);
        runner.setStrictLinting(options.strictLinting);
        
        if (!options.documentPath.empty()) {
            runner.setDocumentPath(options.documentPath);
        }
        
        // Add parameters
        for (const auto& param : options.parameters) {
            runner.addParameter(param.first, param.second);
        }
        
        // Execute macro
        auto start = std::chrono::high_resolution_clock::now();
        MacroExecutionInfo result = runner.runMacroFile(options.inputFile);
        auto end = std::chrono::high_resolution_clock::now();
        
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
        
        // Print results
        if (options.verbose) {
            std::cout << "Execution time: " << duration.count() << "ms" << std::endl;
        }
        
        if (options.syntaxCheck) {
            std::cout << "Syntax check results:" << std::endl;
            
            if (result.syntaxResult.isValid) {
                std::cout << "✓ Syntax is valid" << std::endl;
            } else {
                std::cout << "✗ Syntax errors found:" << std::endl;
            }
            
            for (const auto& error : result.syntaxResult.errors) {
                std::cout << "  " << error.severity << " at line " << error.line 
                         << ", column " << error.column << ": " << error.message << std::endl;
            }
            
            if (result.syntaxResult.hasWarnings()) {
                std::cout << "Warnings: " << std::count_if(result.syntaxResult.errors.begin(), 
                                                         result.syntaxResult.errors.end(),
                                                         [](const SyntaxError& e) { 
                                                             return e.severity == "warning"; 
                                                         }) << std::endl;
            }
        }
        
        // Show lint results if available and enabled
        if (!options.disableLinting && !result.lintResult.issues.empty()) {
            if (options.lintOnly) {
                std::cout << "Lint results:" << std::endl;
            } else if (options.verbose) {
                std::cout << "\nLint results:" << std::endl;
            }
            
            if (result.lintResult.isValid) {
                if (options.lintOnly || options.verbose) {
                    std::cout << "✓ No linting errors" << std::endl;
                }
            } else {
                std::cout << "✗ Linting errors found:" << std::endl;
            }
            
            for (const auto& issue : result.lintResult.issues) {
                std::string severityStr;
                switch (issue.severity) {
                    case LintSeverity::Error: severityStr = "error"; break;
                    case LintSeverity::Warning: severityStr = "warning"; break;
                    case LintSeverity::Info: severityStr = "info"; break;
                }
                
                std::cout << "  " << severityStr << " at line " << issue.line;
                if (issue.column > 0) {
                    std::cout << ", column " << issue.column;
                }
                std::cout << " [" << issue.rule << "]: " << issue.message << std::endl;
                
                if (!issue.codeContext.empty() && (options.verbose || options.lintOnly)) {
                    std::cout << "    Code: " << issue.codeContext << std::endl;
                }
            }
            
            if (result.lintResult.hasWarnings() || result.lintResult.hasErrors()) {
                std::cout << "Summary: " << result.lintResult.getErrorCount() << " errors, " 
                         << result.lintResult.getWarningCount() << " warnings" << std::endl;
            }
        }
        
        switch (result.result) {
            case ExecutionResult::Success:
                if (options.verbose) {
                    std::cout << "✓ Macro executed successfully" << std::endl;
                }
                return 0;
                
            case ExecutionResult::SyntaxError:
                std::cerr << "✗ Syntax error: " << result.message << std::endl;
                return 2;
                
            case ExecutionResult::RuntimeError:
                std::cerr << "✗ Runtime error: " << result.message << std::endl;
                return 3;
                
            case ExecutionResult::FileNotFound:
                std::cerr << "✗ File not found: " << options.inputFile << std::endl;
                return 4;
                
            case ExecutionResult::InvalidArguments:
                std::cerr << "✗ Invalid arguments: " << result.message << std::endl;
                return 5;
        }
        
    } catch (const std::exception& e) {
        std::cerr << "Fatal error: " << e.what() << std::endl;
        return 99;
    }
    
    return 0;
}