package protocol

import (
	"testing"

	flatbuffers "github.com/google/flatbuffers/go"
)

// TestRtdOnceGridResultRoundTrip builds an RtdOnceGridResult carrying a
// 2x2 Grid in its value union (guest->host one-shot grid delivery), then
// reads the key and every cell back and asserts equality.
func TestRtdOnceGridResultRoundTrip(t *testing.T) {
	t.Parallel()

	const wantKey = "BDH\x1f AAPL \x1f 30"
	wantCells := [][]float64{{1.0, 2.0}, {3.0, 4.0}}

	b := flatbuffers.NewBuilder(0)

	// Build the 4 Scalar(Num) cells, row-major.
	cellOffsets := make([]flatbuffers.UOffsetT, 0, 4)
	for _, row := range wantCells {
		for _, v := range row {
			NumStart(b)
			NumAddVal(b, v)
			num := NumEnd(b)

			ScalarStart(b)
			ScalarAddValType(b, ScalarValueNum)
			ScalarAddVal(b, num)
			cellOffsets = append(cellOffsets, ScalarEnd(b))
		}
	}

	GridStartDataVector(b, len(cellOffsets))
	for i := len(cellOffsets) - 1; i >= 0; i-- {
		b.PrependUOffsetT(cellOffsets[i])
	}
	dataVec := b.EndVector(len(cellOffsets))

	GridStart(b)
	GridAddRows(b, 2)
	GridAddCols(b, 2)
	GridAddData(b, dataVec)
	grid := GridEnd(b)

	AnyStart(b)
	AnyAddValType(b, AnyValueGrid)
	AnyAddVal(b, grid)
	anyVal := AnyEnd(b)

	keyOff := b.CreateString(wantKey)

	RtdOnceGridResultStart(b)
	RtdOnceGridResultAddKey(b, keyOff)
	RtdOnceGridResultAddValue(b, anyVal)
	root := RtdOnceGridResultEnd(b)
	b.Finish(root)

	buf := b.FinishedBytes()
	got := GetRootAsRtdOnceGridResult(buf, 0)

	if string(got.Key()) != wantKey {
		t.Fatalf("Key mismatch: want %q, got %q", wantKey, string(got.Key()))
	}

	anyOut := new(Any)
	if got.Value(anyOut) == nil {
		t.Fatal("Value() returned nil Any")
	}
	if anyOut.ValType() != AnyValueGrid {
		t.Fatalf("expected AnyValueGrid, got %d", anyOut.ValType())
	}

	gridOut := new(Grid)
	if !anyOut.Val(&gridOut._tab) {
		t.Fatal("failed to read Grid from Any union")
	}
	if gridOut.Rows() != 2 || gridOut.Cols() != 2 {
		t.Fatalf("expected 2x2 grid, got %dx%d", gridOut.Rows(), gridOut.Cols())
	}
	if err := gridOut.Validate(); err != nil {
		t.Fatalf("grid validate: %v", err)
	}
	if gridOut.DataLength() != 4 {
		t.Fatalf("expected 4 cells, got %d", gridOut.DataLength())
	}

	for idx := 0; idx < gridOut.DataLength(); idx++ {
		sc := new(Scalar)
		if !gridOut.Data(sc, idx) {
			t.Fatalf("failed to read cell %d", idx)
		}
		if sc.ValType() != ScalarValueNum {
			t.Fatalf("cell %d: expected ScalarValueNum, got %d", idx, sc.ValType())
		}
		num := new(Num)
		if !sc.Val(&num._tab) {
			t.Fatalf("cell %d: failed to read Num", idx)
		}
		wantV := wantCells[idx/2][idx%2]
		if num.Val() != wantV {
			t.Fatalf("cell %d: want %v, got %v", idx, wantV, num.Val())
		}
	}
}
