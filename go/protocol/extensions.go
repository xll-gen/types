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

// validateDims checks that rows*cols is non-negative, fits in int32
// (Excel/C++ uses int32 for the total count), and matches the data
// vector length. Shared by Grid.Validate and NumGrid.Validate.
func validateDims(rows, cols int32, dataLen int) error {
	if rows < 0 || cols < 0 {
		return fmt.Errorf("negative dimensions: %d x %d", rows, cols)
	}

	expectedCount := uint64(rows) * uint64(cols)

	if expectedCount > math.MaxInt32 {
		return fmt.Errorf("%w: count %d > MaxInt32", ErrOverflow, expectedCount)
	}

	if uint64(dataLen) != expectedCount {
		return fmt.Errorf("%w: expected %d, got %d", ErrInvalidDimensions, expectedCount, dataLen)
	}
	return nil
}

// Validate checks if the Grid dimensions match the data length.
func (rcv *Grid) Validate() error {
	return validateDims(rcv.Rows(), rcv.Cols(), rcv.DataLength())
}

// Validate checks if the NumGrid dimensions match the data length.
func (rcv *NumGrid) Validate() error {
	return validateDims(rcv.Rows(), rcv.Cols(), rcv.DataLength())
}

// Validate checks if the Range has valid references.
func (rcv *Range) Validate() error {
	if rcv.RefsLength() > 65535 {
		return fmt.Errorf("%w: got %d", ErrTooManyRefs, rcv.RefsLength())
	}
	return nil
}
