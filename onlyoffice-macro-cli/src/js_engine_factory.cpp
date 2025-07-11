#include "js_engine_interface.h"

#ifdef HAVE_V8
#include "v8_engine.h"
#endif

#ifdef HAVE_JSC
#include "jsc_engine.h"
#endif

namespace onlyoffice {
namespace macro {

std::unique_ptr<JSEngineInterface> JSEngineFactory::createEngine(EngineType type) {
    switch (type) {
#ifdef HAVE_V8
        case EngineType::V8:
            return std::make_unique<V8Engine>();
#endif
#ifdef HAVE_JSC
        case EngineType::JavaScriptCore:
            return std::make_unique<JSCEngine>();
#endif
        case EngineType::Auto:
            return createEngine(getBestAvailableEngine());
        default:
            return nullptr;
    }
}

JSEngineFactory::EngineType JSEngineFactory::getBestAvailableEngine() {
#ifdef __APPLE__
    #ifdef HAVE_JSC
        return EngineType::JavaScriptCore;
    #elif defined(HAVE_V8)
        return EngineType::V8;
    #endif
#else
    #ifdef HAVE_V8
        return EngineType::V8;
    #elif defined(HAVE_JSC)
        return EngineType::JavaScriptCore;
    #endif
#endif
    
    // No engine available
    return EngineType::Auto;
}

bool JSEngineFactory::isEngineAvailable(EngineType type) {
    switch (type) {
        case EngineType::V8:
#ifdef HAVE_V8
            return true;
#else
            return false;
#endif
        case EngineType::JavaScriptCore:
#ifdef HAVE_JSC
            return true;
#else
            return false;
#endif
        case EngineType::Auto:
            return isEngineAvailable(getBestAvailableEngine());
        default:
            return false;
    }
}

std::string JSEngineFactory::getEngineTypeName(EngineType type) {
    switch (type) {
        case EngineType::V8:
            return "V8";
        case EngineType::JavaScriptCore:
            return "JavaScriptCore";
        case EngineType::Auto:
            return "Auto";
        default:
            return "Unknown";
    }
}

} // namespace macro
} // namespace onlyoffice