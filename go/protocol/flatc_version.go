package protocol

// FlatcVersion records the flatc compiler version that produced the
// generated Go files in this package (the *_generated.go files).
//
// Source of truth: xll-gen's `internal/versions.FlatBuffers` constant.
// When regenerating, update BOTH this constant AND xll-gen in the same
// release cycle — they're a co-change cluster.
//
// xll-gen's `cmd/doctor` performs a runtime check that this value
// matches its own pinned version; a mismatch indicates the types
// module was bumped without a matching flatc rerun on the generated
// files (or vice versa), which can produce subtle wire-incompat bugs.
//
// To regenerate (in this repo):
//
//	flatc --go --go-namespace protocol --go-module-name <none> \
//	      -o go/protocol/ go/protocol/protocol.fbs
//
// then update FlatcVersion to the version string `flatc --version`
// reports.
const FlatcVersion = "25.9.23"
