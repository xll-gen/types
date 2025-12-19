package protocol

import (
	flatbuffers "github.com/google/flatbuffers/go"
	"testing"
)

func TestGrid_Validate(t *testing.T) {
	t.Parallel()
	b := flatbuffers.NewBuilder(0)

	// Helper to create Scalar Int
	createScalarInt := func(val int32) flatbuffers.UOffsetT {
		IntStart(b)
		IntAddVal(b, val)
		intOff := IntEnd(b)

		ScalarStart(b)
		ScalarAddValType(b, ScalarValueInt)
		ScalarAddVal(b, intOff)
		return ScalarEnd(b)
	}

	// Valid Grid 2x2
	s1 := createScalarInt(1)
	s2 := createScalarInt(2)
	s3 := createScalarInt(3)
	s4 := createScalarInt(4)

	GridStartDataVector(b, 4)
	b.PrependUOffsetT(s4)
	b.PrependUOffsetT(s3)
	b.PrependUOffsetT(s2)
	b.PrependUOffsetT(s1)
	data := b.EndVector(4)

	GridStart(b)
	GridAddRows(b, 2)
	GridAddCols(b, 2)
	GridAddData(b, data)
	gridOff := GridEnd(b)
	b.Finish(gridOff)

	buf := b.FinishedBytes()
	grid := GetRootAsGrid(buf, 0)

	if err := grid.Validate(); err != nil {
		t.Errorf("expected valid grid, got error: %v", err)
	}

	// Invalid Grid (2x2 but 2 elements)
	b.Reset()
	s1 = createScalarInt(1)
	s2 = createScalarInt(2)

	GridStartDataVector(b, 2)
	b.PrependUOffsetT(s2)
	b.PrependUOffsetT(s1)
	data = b.EndVector(2)

	GridStart(b)
	GridAddRows(b, 2)
	GridAddCols(b, 2)
	GridAddData(b, data)
	gridOff = GridEnd(b)
	b.Finish(gridOff)

	buf = b.FinishedBytes()
	grid = GetRootAsGrid(buf, 0)

	if err := grid.Validate(); err == nil {
		t.Error("expected error for invalid grid, got nil")
	}
}
