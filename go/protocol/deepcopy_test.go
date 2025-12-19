package protocol

import (
	flatbuffers "github.com/google/flatbuffers/go"
	"testing"
)

func TestScalar_Clone(t *testing.T) {
	t.Parallel()
	b := flatbuffers.NewBuilder(0)

	// Create Scalar Int 42
	IntStart(b)
	IntAddVal(b, 42)
	intOff := IntEnd(b)
	ScalarStart(b)
	ScalarAddValType(b, ScalarValueInt)
	ScalarAddVal(b, intOff)
	b.Finish(ScalarEnd(b))

	orig := GetRootAsScalar(b.FinishedBytes(), 0)
	clone := orig.Clone()

	// Verify deep copy by checking buffer independence (simple address check isn't always robust due to GC/Stack, but slices pointing to different arrays is key)
	// We can check if modification to one buffer affects the other, but these are read-only.
	// Instead, we just check values.

	if clone.ValType() != ScalarValueInt {
		t.Errorf("Expected type Int, got %v", clone.ValType())
	}

	var i Int
	if !clone.Val(&i._tab) {
		t.Error("Failed to access value")
	}
	if i.Val() != 42 {
		t.Errorf("Expected 42, got %d", i.Val())
	}

	// Create Scalar String
	b.Reset()
	strOff := b.CreateString("hello")
	StrStart(b)
	StrAddVal(b, strOff)
	sOff := StrEnd(b)
	ScalarStart(b)
	ScalarAddValType(b, ScalarValueStr)
	ScalarAddVal(b, sOff)
	b.Finish(ScalarEnd(b))

	origStr := GetRootAsScalar(b.FinishedBytes(), 0)
	cloneStr := origStr.Clone()

	if cloneStr.ValType() != ScalarValueStr {
		t.Errorf("Expected type Str, got %v", cloneStr.ValType())
	}
	var s Str
	if !cloneStr.Val(&s._tab) {
		t.Error("Failed to access value")
	}
	if string(s.Val()) != "hello" {
		t.Errorf("Expected 'hello', got '%s'", s.Val())
	}
}

func TestGrid_Clone(t *testing.T) {
	t.Parallel()
	b := flatbuffers.NewBuilder(0)

	// Create 2x1 Grid with [42, 100]
	// Scalar 1
	IntStart(b)
	IntAddVal(b, 42)
	i1 := IntEnd(b)
	ScalarStart(b)
	ScalarAddValType(b, ScalarValueInt)
	ScalarAddVal(b, i1)
	s1 := ScalarEnd(b)

	// Scalar 2
	NumStart(b)
	NumAddVal(b, 100.5)
	n1 := NumEnd(b)
	ScalarStart(b)
	ScalarAddValType(b, ScalarValueNum)
	ScalarAddVal(b, n1)
	s2 := ScalarEnd(b)

	GridStartDataVector(b, 2)
	b.PrependUOffsetT(s2)
	b.PrependUOffsetT(s1)
	data := b.EndVector(2)

	GridStart(b)
	GridAddRows(b, 2)
	GridAddCols(b, 1)
	GridAddData(b, data)
	b.Finish(GridEnd(b))

	orig := GetRootAsGrid(b.FinishedBytes(), 0)
	clone := orig.Clone()

	if clone.Rows() != 2 || clone.Cols() != 1 {
		t.Errorf("Dimensions mismatch: %dx%d", clone.Rows(), clone.Cols())
	}
	if clone.DataLength() != 2 {
		t.Errorf("Data length mismatch: %d", clone.DataLength())
	}

	// Check S1
	var sc Scalar
	if !clone.Data(&sc, 0) {
		t.Error("Failed to get data 0")
	}
	if sc.ValType() != ScalarValueInt {
		t.Error("Expected Int at 0")
	}
	var iv Int
	sc.Val(&iv._tab)
	if iv.Val() != 42 {
		t.Errorf("Expected 42, got %d", iv.Val())
	}

	// Check S2
	if !clone.Data(&sc, 1) {
		t.Error("Failed to get data 1")
	}
	if sc.ValType() != ScalarValueNum {
		t.Error("Expected Num at 1")
	}
	var nv Num
	sc.Val(&nv._tab)
	if nv.Val() != 100.5 {
		t.Errorf("Expected 100.5, got %f", nv.Val())
	}
}

func TestRange_Clone(t *testing.T) {
	t.Parallel()
	b := flatbuffers.NewBuilder(0)

	name := b.CreateString("Sheet1")

	RangeStartRefsVector(b, 1)
	CreateRect(b, 1, 2, 3, 4)
	refs := b.EndVector(1)

	RangeStart(b)
	RangeAddSheetName(b, name)
	RangeAddRefs(b, refs)
	b.Finish(RangeEnd(b))

	orig := GetRootAsRange(b.FinishedBytes(), 0)
	clone := orig.Clone()

	if string(clone.SheetName()) != "Sheet1" {
		t.Errorf("Expected Sheet1, got %s", clone.SheetName())
	}
	if clone.RefsLength() != 1 {
		t.Errorf("Expected 1 ref, got %d", clone.RefsLength())
	}
	var r Rect
	if !clone.Refs(&r, 0) {
		t.Error("Failed to get ref")
	}
	if r.RowFirst() != 1 || r.RowLast() != 2 || r.ColFirst() != 3 || r.ColLast() != 4 {
		t.Error("Rect mismatch")
	}
}

func TestAny_Clone(t *testing.T) {
	t.Parallel()
	b := flatbuffers.NewBuilder(0)

	// Create Any containing Range
	name := b.CreateString("Sheet1")
	RangeStartRefsVector(b, 0)
	refs := b.EndVector(0)
	RangeStart(b)
	RangeAddSheetName(b, name)
	RangeAddRefs(b, refs)
	rOff := RangeEnd(b)

	AnyStart(b)
	AnyAddValType(b, AnyValueRange)
	AnyAddVal(b, rOff)
	b.Finish(AnyEnd(b))

	orig := GetRootAsAny(b.FinishedBytes(), 0)
	clone := orig.Clone()

	if clone.ValType() != AnyValueRange {
		t.Error("Expected Range type")
	}
	var r Range
	if !clone.Val(&r._tab) {
		t.Error("Failed to get Range")
	}
	if string(r.SheetName()) != "Sheet1" {
		t.Error("Sheet name mismatch")
	}
}
