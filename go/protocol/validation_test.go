package protocol

import (
	"errors"
	"math"
	"testing"

	flatbuffers "github.com/google/flatbuffers/go"
)

func TestRangeValidation(t *testing.T) {
	t.Parallel()
	// Case 1: Range with > 65535 refs
	b := flatbuffers.NewBuilder(0)

	count := 65536
	RangeStartRefsVector(b, count)
	for i := 0; i < count; i++ {
		CreateRect(b, 0, 0, 0, 0)
	}
	refs := b.EndVector(count)

	RangeStart(b)
	RangeAddRefs(b, refs)
	off := RangeEnd(b)
	b.Finish(off)

	buf := b.FinishedBytes()
	r := GetRootAsRange(buf, 0)

	err := r.Validate()
	if err == nil {
		t.Error("Expected error for Range with 65536 refs, got nil")
	}
	if !errors.Is(err, ErrTooManyRefs) {
		t.Errorf("Expected ErrTooManyRefs, got %v", err)
	}
}

func TestGridOverflowValidation(t *testing.T) {
	t.Parallel()
	// Case: Rows * Cols > MaxInt32
	// Rows = MaxInt32, Cols = 2

	b := flatbuffers.NewBuilder(0)
	GridStartDataVector(b, 0)
	data := b.EndVector(0)

	GridStart(b)
	GridAddRows(b, math.MaxInt32)
	GridAddCols(b, 2)
	GridAddData(b, data)
	off := GridEnd(b)
	b.Finish(off)

	buf := b.FinishedBytes()
	g := GetRootAsGrid(buf, 0)

	err := g.Validate()
	if err == nil {
		t.Fatal("Expected error, got nil")
	}

	if !errors.Is(err, ErrOverflow) {
		t.Errorf("Expected ErrOverflow, got %v", err)
	}
}
