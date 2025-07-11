#pragma once

#include <string>
#include <vector>
#include <map>

namespace onlyoffice {
namespace macro {

struct CLIOptions {
    std::string inputFile;
    std::string outputFile;
    std::string documentPath;
    bool syntaxCheck = false;
    bool verbose = false;
    bool help = false;
    bool version = false;
    bool dryRun = false;
    bool lintOnly = false;
    bool disableLinting = false;
    bool strictLinting = false;
    std::map<std::string, std::string> parameters;
};

class CLIParser {
public:
    CLIParser();
    ~CLIParser();
    
    CLIOptions parse(int argc, char* argv[]);
    void printHelp() const;
    void printVersion() const;
    
private:
    void setupOptions();
    std::string getUsageString() const;
    std::string getHelpString() const;
};

} // namespace macro
} // namespace onlyoffice