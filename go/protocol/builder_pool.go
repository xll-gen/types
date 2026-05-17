package protocol

import (
	"sync"

	flatbuffers "github.com/google/flatbuffers/go"
)

// builderPool reuses *flatbuffers.Builder instances across deepcopy
// (Clone) calls. The pre-v0.2.5 implementation allocated a fresh
// builder per Clone via flatbuffers.NewBuilder(0); under xll-gen's
// async batcher and RTD handler hot paths that was the dominant
// allocation in steady state.
//
// Pool semantics: builders are sized to 1024 bytes on first allocation;
// Reset() preserves the underlying buffer for the next caller, so
// sustained workloads converge to a single per-goroutine builder live
// in the pool. Callers MUST call releaseBuilder before discarding to
// avoid leaking builders that the GC eventually frees but never recycles.
var builderPool = sync.Pool{
	New: func() any { return flatbuffers.NewBuilder(1024) },
}

// acquireBuilder returns a builder from the pool. Always pair with
// releaseBuilder via defer.
func acquireBuilder() *flatbuffers.Builder {
	return builderPool.Get().(*flatbuffers.Builder)
}

// releaseBuilder resets the builder's offset cursor and returns it to
// the pool. Resetting (not freeing) means the backing buffer is
// retained — the next caller skips the allocation as long as their
// serialized output fits in the prior peak size.
func releaseBuilder(b *flatbuffers.Builder) {
	b.Reset()
	builderPool.Put(b)
}
