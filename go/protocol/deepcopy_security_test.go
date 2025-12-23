package protocol

import (
	flatbuffers "github.com/google/flatbuffers/go"
	"testing"
	"encoding/binary"
	"math"
)

func TestGrid_DeepCopy_DoS(t *testing.T) {
	b := flatbuffers.NewBuilder(0)

	// Create a small Grid with 1 element
	GridStartDataVector(b, 1)
	b.PrependUOffsetT(0)
	data := b.EndVector(1)

	GridStart(b)
	GridAddData(b, data)
	b.Finish(GridEnd(b))

	buf := b.FinishedBytes()

	// Locate the vector length (1) and change it to MaxInt32 + 100
	found := false
	target := uint32(1)
	replacement := uint32(math.MaxInt32) + 100

	// Scan from end backwards
	for i := len(buf)-4; i >= 0; i-- {
		if binary.LittleEndian.Uint32(buf[i:]) == target {
			// Found it. We assume it's the vector length.
			binary.LittleEndian.PutUint32(buf[i:], replacement)
			found = true
			break
		}
	}

	if !found {
		t.Fatal("Could not locate vector length in buffer")
	}

	g := GetRootAsGrid(buf, 0)

	defer func() {
		if r := recover(); r != nil {
			t.Fatalf("DoS Reproduction: Panic occurred: %v", r)
		}
	}()

	off := g.DeepCopy(flatbuffers.NewBuilder(0))

	// If fix is working, it should return 0 (graceful failure)
	if off != 0 {
		t.Errorf("DeepCopy returned non-zero offset %d for invalid length, expected 0", off)
	}
}
