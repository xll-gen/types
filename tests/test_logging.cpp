#include "types/utility.h"
#include <cassert>
#include <iostream>

#ifdef _WIN32
#include <windows.h>
HINSTANCE g_hModule = NULL;
#endif

void TestLogging() {
    // Default should be false
    assert(GetDebugFlag() == false);

    // Enable
    SetDebugFlag(true);
    assert(GetDebugFlag() == true);

    // Log something (should not crash)
    DebugLog("Testing debug log: %d\n", 123);

    // Disable
    SetDebugFlag(false);
    assert(GetDebugFlag() == false);

    // Log something (should be ignored)
    DebugLog("This should not appear.\n");

    std::cout << "Logging test passed." << std::endl;
}

int main() {
    TestLogging();
    return 0;
}
