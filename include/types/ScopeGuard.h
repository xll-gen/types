#ifndef SCOPE_GUARD_H
#define SCOPE_GUARD_H

#include <functional>

// A simple ScopeGuard for RAII cleanup
class ScopeGuard {
public:
    explicit ScopeGuard(std::function<void()> onExit)
        : onExit_(onExit), dismissed_(false) {}

    ~ScopeGuard() {
        if (!dismissed_) {
            onExit_();
        }
    }

    void Dismiss() {
        dismissed_ = true;
    }

private:
    std::function<void()> onExit_;
    bool dismissed_;
};

#endif // SCOPE_GUARD_H
