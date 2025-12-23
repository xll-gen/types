package protocol

import (
	flatbuffers "github.com/google/flatbuffers/go"
	"math"
)

// Clone creates a deep copy of the Scalar.
func (rcv *Scalar) Clone() *Scalar {
	if rcv == nil {
		return nil
	}
	b := flatbuffers.NewBuilder(0)
	off := rcv.DeepCopy(b)
	b.Finish(off)
	buf := b.FinishedBytes()
	// Create a new buffer with exact size to own the data
	newBuf := make([]byte, len(buf))
	copy(newBuf, buf)
	return GetRootAsScalar(newBuf, 0)
}

// DeepCopy serializes the Scalar into the builder.
func (rcv *Scalar) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	if rcv == nil {
		return 0
	}

	valType := rcv.ValType()
	var valOffset flatbuffers.UOffsetT

	switch valType {
	case ScalarValueBool:
		t := new(Bool)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case ScalarValueNum:
		t := new(Num)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case ScalarValueInt:
		t := new(Int)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case ScalarValueStr:
		t := new(Str)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case ScalarValueErr:
		t := new(Err)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case ScalarValueAsyncHandle:
		t := new(AsyncHandle)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case ScalarValueNil:
		t := new(Nil)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	}

	ScalarStart(b)
	if valOffset != 0 {
		ScalarAddValType(b, valType)
		ScalarAddVal(b, valOffset)
	}
	return ScalarEnd(b)
}

// DeepCopy helpers for ScalarValue types

func (rcv *Bool) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	BoolStart(b)
	BoolAddVal(b, rcv.Val())
	return BoolEnd(b)
}

func (rcv *Num) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	NumStart(b)
	NumAddVal(b, rcv.Val())
	return NumEnd(b)
}

func (rcv *Int) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	IntStart(b)
	IntAddVal(b, rcv.Val())
	return IntEnd(b)
}

func (rcv *Str) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	s := rcv.Val()
	off := b.CreateString(string(s))
	StrStart(b)
	StrAddVal(b, off)
	return StrEnd(b)
}

func (rcv *Err) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	ErrStart(b)
	ErrAddVal(b, rcv.Val())
	return ErrEnd(b)
}

func (rcv *AsyncHandle) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	l := rcv.ValLength()
	if l > math.MaxInt32 {
		return 0
	}
	AsyncHandleStartValVector(b, l)
	for i := l - 1; i >= 0; i-- {
		b.PrependByte(rcv.Val(i))
	}
	vec := b.EndVector(l)
	AsyncHandleStart(b)
	AsyncHandleAddVal(b, vec)
	return AsyncHandleEnd(b)
}

func (rcv *Nil) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	NilStart(b)
	return NilEnd(b)
}

// Clone creates a deep copy of the Grid.
func (rcv *Grid) Clone() *Grid {
	if rcv == nil {
		return nil
	}
	b := flatbuffers.NewBuilder(0)
	off := rcv.DeepCopy(b)
	b.Finish(off)
	buf := b.FinishedBytes()
	newBuf := make([]byte, len(buf))
	copy(newBuf, buf)
	return GetRootAsGrid(newBuf, 0)
}

// DeepCopy serializes the Grid into the builder.
func (rcv *Grid) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	if rcv == nil {
		return 0
	}

	l := rcv.DataLength()
	if l < 0 || l > math.MaxInt32 {
		return 0
	}

	offsets := make([]flatbuffers.UOffsetT, l)
	s := new(Scalar)
	for i := 0; i < l; i++ {
		if rcv.Data(s, i) {
			offsets[i] = s.DeepCopy(b)
		}
	}

	GridStartDataVector(b, l)
	for i := l - 1; i >= 0; i-- {
		b.PrependUOffsetT(offsets[i])
	}
	dataOff := b.EndVector(l)

	GridStart(b)
	GridAddRows(b, rcv.Rows())
	GridAddCols(b, rcv.Cols())
	GridAddData(b, dataOff)
	return GridEnd(b)
}

// Clone creates a deep copy of the NumGrid.
func (rcv *NumGrid) Clone() *NumGrid {
	if rcv == nil {
		return nil
	}
	b := flatbuffers.NewBuilder(0)
	off := rcv.DeepCopy(b)
	b.Finish(off)
	buf := b.FinishedBytes()
	newBuf := make([]byte, len(buf))
	copy(newBuf, buf)
	return GetRootAsNumGrid(newBuf, 0)
}

// DeepCopy serializes the NumGrid into the builder.
func (rcv *NumGrid) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	if rcv == nil {
		return 0
	}

	l := rcv.DataLength()
	if l < 0 || l > math.MaxInt32 {
		return 0
	}

	NumGridStartDataVector(b, l)
	for i := l - 1; i >= 0; i-- {
		b.PrependFloat64(rcv.Data(i))
	}
	dataOff := b.EndVector(l)

	NumGridStart(b)
	NumGridAddRows(b, rcv.Rows())
	NumGridAddCols(b, rcv.Cols())
	NumGridAddData(b, dataOff)
	return NumGridEnd(b)
}

// Clone creates a deep copy of the Range.
func (rcv *Range) Clone() *Range {
	if rcv == nil {
		return nil
	}
	b := flatbuffers.NewBuilder(0)
	off := rcv.DeepCopy(b)
	b.Finish(off)
	buf := b.FinishedBytes()
	newBuf := make([]byte, len(buf))
	copy(newBuf, buf)
	return GetRootAsRange(newBuf, 0)
}

// DeepCopy serializes the Range into the builder.
func (rcv *Range) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	if rcv == nil {
		return 0
	}

	// Strings
	sheetName := rcv.SheetName()
	var sheetNameOff flatbuffers.UOffsetT
	if sheetName != nil {
		sheetNameOff = b.CreateString(string(sheetName))
	}

	format := rcv.Format()
	var formatOff flatbuffers.UOffsetT
	if format != nil {
		formatOff = b.CreateString(string(format))
	}

	// Refs Vector of Structs
	l := rcv.RefsLength()
	if l < 0 || l > math.MaxInt32 {
		return 0
	}

	RangeStartRefsVector(b, l)
	r := new(Rect)
	for i := l - 1; i >= 0; i-- {
		if rcv.Refs(r, i) {
			CreateRect(b, r.RowFirst(), r.RowLast(), r.ColFirst(), r.ColLast())
		}
	}
	refsOff := b.EndVector(l)

	RangeStart(b)
	if sheetNameOff != 0 {
		RangeAddSheetName(b, sheetNameOff)
	}
	RangeAddRefs(b, refsOff)
	if formatOff != 0 {
		RangeAddFormat(b, formatOff)
	}
	return RangeEnd(b)
}

// Clone creates a deep copy of the Any.
func (rcv *Any) Clone() *Any {
	if rcv == nil {
		return nil
	}
	b := flatbuffers.NewBuilder(0)
	off := rcv.DeepCopy(b)
	b.Finish(off)
	buf := b.FinishedBytes()
	newBuf := make([]byte, len(buf))
	copy(newBuf, buf)
	return GetRootAsAny(newBuf, 0)
}

// DeepCopy serializes the Any into the builder.
func (rcv *Any) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	if rcv == nil {
		return 0
	}

	valType := rcv.ValType()
	var valOffset flatbuffers.UOffsetT

	switch valType {
	case AnyValueBool:
		t := new(Bool)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueNum:
		t := new(Num)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueInt:
		t := new(Int)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueStr:
		t := new(Str)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueErr:
		t := new(Err)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueAsyncHandle:
		t := new(AsyncHandle)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueNil:
		t := new(Nil)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueGrid:
		t := new(Grid)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueNumGrid:
		t := new(NumGrid)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueRange:
		t := new(Range)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	case AnyValueRefCache:
		t := new(RefCache)
		if rcv.Val(&t._tab) {
			valOffset = t.DeepCopy(b)
		}
	}

	AnyStart(b)
	if valOffset != 0 {
		AnyAddValType(b, valType)
		AnyAddVal(b, valOffset)
	}
	return AnyEnd(b)
}

func (rcv *RefCache) DeepCopy(b *flatbuffers.Builder) flatbuffers.UOffsetT {
	k := rcv.Key()
	off := b.CreateString(string(k))
	RefCacheStart(b)
	RefCacheAddKey(b, off)
	return RefCacheEnd(b)
}
