#pragma once

#include <string>
#include <vector>
#include <map>
#include <memory>
#include "linter.h"

namespace onlyoffice {
namespace macro {

/**
 * @brief JavaScript Engine Interface
 * 
 * Abstract base class for JavaScript engines used in macro validation and execution.
 * Provides a unified interface for different JavaScript engines (V8, JavaScriptCore).
 */
class JSEngineInterface {
public:
    virtual ~JSEngineInterface() = default;
    
    /**
     * @brief Initialize the JavaScript engine
     * @return true if initialization successful, false otherwise
     */
    virtual bool initialize() = 0;
    
    /**
     * @brief Cleanup engine resources
     */
    virtual void cleanup() = 0;
    
    /**
     * @brief Validate JavaScript syntax
     * @param source JavaScript source code
     * @param issues Vector to store validation issues
     * @return true if syntax is valid, false otherwise
     */
    virtual bool validateSyntax(const std::string& source, std::vector<LintIssue>& issues) = 0;
    
    /**
     * @brief Execute JavaScript code
     * @param source JavaScript source code
     * @param params Parameters to pass to script
     * @return true if execution successful, false otherwise
     */
    virtual bool executeScript(const std::string& source, const std::map<std::string, std::string>& params = {}) = 0;
    
    /**
     * @brief Setup OnlyOffice API mock objects
     */
    virtual void setupOnlyOfficeAPI() = 0;
    
    /**
     * @brief Get last error message
     * @return Error message string
     */
    virtual std::string getLastError() const = 0;
    
    /**
     * @brief Clear stored errors
     */
    virtual void clearErrors() = 0;
    
    /**
     * @brief Get engine name
     * @return Engine name string
     */
    virtual std::string getEngineName() const = 0;
    
    /**
     * @brief Get engine version
     * @return Engine version string
     */
    virtual std::string getEngineVersion() const = 0;
    
    /**
     * @brief Check if engine is initialized
     * @return true if initialized, false otherwise
     */
    virtual bool isInitialized() const = 0;
};

/**
 * @brief JavaScript Engine Factory
 * 
 * Factory class for creating JavaScript engine instances.
 */
class JSEngineFactory {
public:
    enum class EngineType {
        V8,                 ///< V8 JavaScript engine
        JavaScriptCore,     ///< JavaScriptCore engine (Apple platforms)
        Auto               ///< Automatically select best available engine
    };
    
    /**
     * @brief Create JavaScript engine instance
     * @param type Engine type to create
     * @return Unique pointer to engine instance
     */
    static std::unique_ptr<JSEngineInterface> createEngine(EngineType type);
    
    /**
     * @brief Get best available engine type for current platform
     * @return Best available engine type
     */
    static EngineType getBestAvailableEngine();
    
    /**
     * @brief Check if engine type is available
     * @param type Engine type to check
     * @return true if available, false otherwise
     */
    static bool isEngineAvailable(EngineType type);
    
    /**
     * @brief Get engine type name
     * @param type Engine type
     * @return Engine type name string
     */
    static std::string getEngineTypeName(EngineType type);
};

} // namespace macro
} // namespace onlyoffice