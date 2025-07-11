#pragma once

#include <set>
#include <string>

namespace onlyoffice {
namespace macro {

void setupKnownAPIs(std::set<std::string>& knownAPIs);
void setupKnownGlobals(std::set<std::string>& knownGlobals);

} // namespace macro
} // namespace onlyoffice