#include "cli_parser.h"
#include <iostream>
#include <algorithm>
#include <sstream>

namespace onlyoffice {
namespace macro {

CLIParser::CLIParser() {
    setupOptions();
}

CLIParser::~CLIParser() = default;

CLIOptions CLIParser::parse(int argc, char* argv[]) {
    CLIOptions options;
    
    for (int i = 1; i < argc; ++i) {
        std::string arg = argv[i];
        
        if (arg == "-h" || arg == "--help") {
            options.help = true;
        } else if (arg == "-v" || arg == "--version") {
            options.version = true;
        } else if (arg == "-c" || arg == "--syntax-check") {
            options.syntaxCheck = true;
        } else if (arg == "--verbose") {
            options.verbose = true;
        } else if (arg == "--dry-run") {
            options.dryRun = true;
        } else if (arg == "-l" || arg == "--lint") {
            options.lintOnly = true;
        } else if (arg == "--no-lint") {
            options.disableLinting = true;
        } else if (arg == "--strict-lint") {
            options.strictLinting = true;
        } else if (arg == "-o" || arg == "--output") {
            if (i + 1 < argc) {
                options.outputFile = argv[++i];
            } else {
                throw std::runtime_error("Missing argument for " + arg);
            }
        } else if (arg == "-d" || arg == "--document") {
            if (i + 1 < argc) {
                options.documentPath = argv[++i];
            } else {
                throw std::runtime_error("Missing argument for " + arg);
            }
        } else if (arg == "-p" || arg == "--param") {
            if (i + 1 < argc) {
                std::string param = argv[++i];
                size_t pos = param.find('=');
                if (pos != std::string::npos) {
                    std::string key = param.substr(0, pos);
                    std::string value = param.substr(pos + 1);
                    options.parameters[key] = value;
                } else {
                    throw std::runtime_error("Invalid parameter format: " + param + " (expected key=value)");
                }
            } else {
                throw std::runtime_error("Missing argument for " + arg);
            }
        } else if (arg.front() == '-') {
            throw std::runtime_error("Unknown option: " + arg);
        } else {
            if (options.inputFile.empty()) {
                options.inputFile = arg;
            } else {
                throw std::runtime_error("Multiple input files not supported");
            }
        }
    }
    
    return options;
}

void CLIParser::printHelp() const {
    std::cout << getUsageString() << std::endl;
    std::cout << getHelpString() << std::endl;
}

void CLIParser::printVersion() const {
    std::cout << "OnlyOffice Macro CLI v1.0.0" << std::endl;
    std::cout << "JavaScript macro runner with syntax checking" << std::endl;
}

void CLIParser::setupOptions() {
    // Options are handled in parse() method
}

std::string CLIParser::getUsageString() const {
    return "Usage: onlyoffice-macro-cli [OPTIONS] <macro-file.js>";
}

std::string CLIParser::getHelpString() const {
    return R"(
OnlyOffice Macro CLI - JavaScript macro runner with syntax checking

ARGUMENTS:
  <macro-file.js>           JavaScript macro file to execute

OPTIONS:
  -h, --help               Show this help message
  -v, --version            Show version information
  -c, --syntax-check       Only check syntax, don't execute
  -l, --lint               Only run linter, don't execute
  --no-lint                Disable linting
  --strict-lint            Enable strict linting mode
  --verbose                Enable verbose output
  --dry-run                Parse and validate but don't execute
  -o, --output <file>      Output file for results
  -d, --document <path>    Document path for macro context
  -p, --param <key=value>  Set parameter for macro execution

EXAMPLES:
  # Run a macro file
  onlyoffice-macro-cli my-macro.js

  # Check syntax only
  onlyoffice-macro-cli -c my-macro.js

  # Run linter only
  onlyoffice-macro-cli -l my-macro.js

  # Run with strict linting
  onlyoffice-macro-cli --strict-lint my-macro.js

  # Run without linting
  onlyoffice-macro-cli --no-lint my-macro.js

  # Run with verbose output
  onlyoffice-macro-cli --verbose my-macro.js

  # Run with parameters
  onlyoffice-macro-cli -p name=John -p age=30 my-macro.js

  # Run with document context
  onlyoffice-macro-cli -d /path/to/document.docx my-macro.js

  # Dry run (validate only)
  onlyoffice-macro-cli --dry-run my-macro.js
)";
}

} // namespace macro
} // namespace onlyoffice