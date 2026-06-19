package protocol

import "testing"

// TestUnionDeepCopyCompleteness is the Go counterpart of the static_asserts in
// types/src/converters.cpp (R29). The union-tag switches in deepcopy.go
// (AnyValue around the Any DeepCopy, ScalarValue around the Scalar DeepCopy)
// fall through to `return 0` on an unknown tag, so a newly-appended FlatBuffer
// union member would be SILENTLY dropped from a deep copy rather than fail
// loudly. Go has no compile-time switch-exhaustiveness check, so this test
// pins the member count instead.
//
// If this fails, a member was added to ScalarValue/AnyValue in protocol.fbs.
// Update every ladder that dispatches on the union tag — go/protocol/deepcopy.go
// (both switches) AND the C++ converters.cpp ladders (AnyToXLOPER12,
// GridToXLOPER12's per-cell switch, ConvertScalar/ConvertAny) — then bump the
// expected count here and the matching static_asserts in converters.cpp.
func TestUnionDeepCopyCompleteness(t *testing.T) {
	// EnumNames* include the NONE sentinel (value 0), so the counts are
	// member-count + 1: ScalarValue NONE..Date = 9, AnyValue NONE..Date = 13.
	if got := len(EnumNamesScalarValue); got != 9 {
		t.Fatalf("ScalarValue member count = %d, want 9: a member changed in protocol.fbs — "+
			"update deepcopy.go's ScalarValue switch and converters.cpp's ScalarValue ladders, then bump this count", got)
	}
	if got := len(EnumNamesAnyValue); got != 13 {
		t.Fatalf("AnyValue member count = %d, want 13: a member changed in protocol.fbs — "+
			"update deepcopy.go's AnyValue switch and converters.cpp's AnyValue ladders, then bump this count", got)
	}
}
