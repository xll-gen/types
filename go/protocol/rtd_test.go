package protocol

import (
	"testing"

	flatbuffers "github.com/google/flatbuffers/go"
)

func TestRtdConnectRequestDeepCopy(t *testing.T) {
	t.Parallel()

	b := flatbuffers.NewBuilder(0)
	s1 := b.CreateString("Topic1")
	s2 := b.CreateString("Param1")

	RtdConnectRequestStartStringsVector(b, 2)
	b.PrependUOffsetT(s2)
	b.PrependUOffsetT(s1)
	stringsVec := b.EndVector(2)

	RtdConnectRequestStart(b)
	RtdConnectRequestAddTopicId(b, 101)
	RtdConnectRequestAddStrings(b, stringsVec)
	RtdConnectRequestAddNewValues(b, true)
	orc := RtdConnectRequestEnd(b)
	b.Finish(orc)

	buf := b.FinishedBytes()
	req := GetRootAsRtdConnectRequest(buf, 0)

	cloned := req.Clone()

	if cloned.TopicId() != 101 {
		t.Errorf("Expected TopicId 101, got %d", cloned.TopicId())
	}
	if cloned.NewValues() != true {
		t.Errorf("Expected NewValues true, got false")
	}
	if cloned.StringsLength() != 2 {
		t.Errorf("Expected StringsLength 2, got %d", cloned.StringsLength())
	}
	if string(cloned.Strings(0)) != "Topic1" {
		t.Errorf("Expected Strings(0) 'Topic1', got '%s'", string(cloned.Strings(0)))
	}
	if string(cloned.Strings(1)) != "Param1" {
		t.Errorf("Expected Strings(1) 'Param1', got '%s'", string(cloned.Strings(1)))
	}
}

func TestRtdUpdateDeepCopy(t *testing.T) {
	t.Parallel()

	b := flatbuffers.NewBuilder(0)

	// Create Any(Num)
	NumStart(b)
	NumAddVal(b, 123.456)
	n := NumEnd(b)

	AnyStart(b)
	AnyAddValType(b, AnyValueNum)
	AnyAddVal(b, n)
	anyVal := AnyEnd(b)

	RtdUpdateStart(b)
	RtdUpdateAddTopicId(b, 202)
	RtdUpdateAddVal(b, anyVal)
	orc := RtdUpdateEnd(b)
	b.Finish(orc)

	buf := b.FinishedBytes()
	upd := GetRootAsRtdUpdate(buf, 0)

	cloned := upd.Clone()

	if cloned.TopicId() != 202 {
		t.Errorf("Expected TopicId 202, got %d", cloned.TopicId())
	}

	val := new(Any)
	cloned.Val(val)
	if val.ValType() != AnyValueNum {
		t.Errorf("Expected ValType Num, got %d", val.ValType())
	}

	num := new(Num)
	if !val.Val(&num._tab) {
		t.Fatal("Failed to get Num table")
	}
	if num.Val() != 123.456 {
		t.Errorf("Expected Num val 123.456, got %f", num.Val())
	}
}

func TestBatchRtdUpdateDeepCopy(t *testing.T) {
	t.Parallel()

	b := flatbuffers.NewBuilder(0)

	// Update 1
	NumStart(b)
	NumAddVal(b, 1.1)
	n1 := NumEnd(b)
	AnyStart(b)
	AnyAddValType(b, AnyValueNum)
	AnyAddVal(b, n1)
	a1 := AnyEnd(b)
	RtdUpdateStart(b)
	RtdUpdateAddTopicId(b, 1)
	RtdUpdateAddVal(b, a1)
	u1 := RtdUpdateEnd(b)

	// Update 2
	NumStart(b)
	NumAddVal(b, 2.2)
	n2 := NumEnd(b)
	AnyStart(b)
	AnyAddValType(b, AnyValueNum)
	AnyAddVal(b, n2)
	a2 := AnyEnd(b)
	RtdUpdateStart(b)
	RtdUpdateAddTopicId(b, 2)
	RtdUpdateAddVal(b, a2)
	u2 := RtdUpdateEnd(b)

	BatchRtdUpdateStartUpdatesVector(b, 2)
	b.PrependUOffsetT(u2)
	b.PrependUOffsetT(u1)
	vec := b.EndVector(2)

	BatchRtdUpdateStart(b)
	BatchRtdUpdateAddUpdates(b, vec)
	orc := BatchRtdUpdateEnd(b)
	b.Finish(orc)

	buf := b.FinishedBytes()
	batch := GetRootAsBatchRtdUpdate(buf, 0)

	cloned := batch.Clone()

	if cloned.UpdatesLength() != 2 {
		t.Errorf("Expected UpdatesLength 2, got %d", cloned.UpdatesLength())
	}

	u := new(RtdUpdate)

	// Check first update
	if !cloned.Updates(u, 0) {
		t.Fatal("Failed to get update 0")
	}
	if u.TopicId() != 1 {
		t.Errorf("Expected update 0 TopicId 1, got %d", u.TopicId())
	}

	// Check second update
	if !cloned.Updates(u, 1) {
		t.Fatal("Failed to get update 1")
	}
	if u.TopicId() != 2 {
		t.Errorf("Expected update 1 TopicId 2, got %d", u.TopicId())
	}
}
