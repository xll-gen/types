package protocol

import (
	"errors"
	"fmt"
	"math"
)

var (
	// ErrInvalidDimensions indicates that the grid dimensions do not match the data length.
	ErrInvalidDimensions = errors.New("rows * cols does not match data length")
	// ErrOverflow indicates that the dimensions exceed the maximum supported limits.
	ErrOverflow = errors.New("dimensions exceed maximum supported limits (int32)")
	// ErrTooManyRefs indicates that the range has too many references.
	ErrTooManyRefs = errors.New("too many references (> 65535)")
)

// Validate checks if the Grid dimensions match the data length.
func (rcv *Grid) Validate() error {
	rows := int(rcv.Rows())
	cols := int(rcv.Cols())

	if rows < 0 || cols < 0 {
		return fmt.Errorf("negative dimensions: %d x %d", rows, cols)
	}

	expectedCount := uint64(rcv.Rows()) * uint64(rcv.Cols())

	// Check for int32 overflow as Excel/C++ uses int32 for total count
	if expectedCount > math.MaxInt32 {
		return fmt.Errorf("%w: count %d > MaxInt32", ErrOverflow, expectedCount)
	}

	if uint64(rcv.DataLength()) != expectedCount {
		return fmt.Errorf("%w: expected %d, got %d", ErrInvalidDimensions, expectedCount, rcv.DataLength())
	}
	return nil
}

// Validate checks if the NumGrid dimensions match the data length.
func (rcv *NumGrid) Validate() error {
	rows := int(rcv.Rows())
	cols := int(rcv.Cols())

	if rows < 0 || cols < 0 {
		return fmt.Errorf("negative dimensions: %d x %d", rows, cols)
	}

	expectedCount := uint64(rcv.Rows()) * uint64(rcv.Cols())

	// Check for int32 overflow as Excel/C++ uses int32 for total count
	if expectedCount > math.MaxInt32 {
		return fmt.Errorf("%w: count %d > MaxInt32", ErrOverflow, expectedCount)
	}

	if uint64(rcv.DataLength()) != expectedCount {
		return fmt.Errorf("%w: expected %d, got %d", ErrInvalidDimensions, expectedCount, rcv.DataLength())
	}
	return nil
}

// Validate checks if the Range has valid references.
func (rcv *Range) Validate() error {
	if rcv.RefsLength() > 65535 {
		return fmt.Errorf("%w: got %d", ErrTooManyRefs, rcv.RefsLength())
	}
	return nil
}
