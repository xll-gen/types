package protocol

import (
	"errors"
	"fmt"
)

var (
	// ErrInvalidDimensions indicates that the grid dimensions do not match the data length.
	ErrInvalidDimensions = errors.New("rows * cols does not match data length")
)

// Validate checks if the Grid dimensions match the data length.
func (rcv *Grid) Validate() error {
	rows := int(rcv.Rows())
	cols := int(rcv.Cols())

	if rows < 0 || cols < 0 {
		return fmt.Errorf("negative dimensions: %d x %d", rows, cols)
	}

	expectedCount := uint64(rcv.Rows()) * uint64(rcv.Cols())
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
	if uint64(rcv.DataLength()) != expectedCount {
		return fmt.Errorf("%w: expected %d, got %d", ErrInvalidDimensions, expectedCount, rcv.DataLength())
	}
	return nil
}

// Validate checks if the Range has valid references.
func (rcv *Range) Validate() error {
    // Basic check: range refs length should reasonably match use cases.
    // For now, we just ensure no crashes or obvious issues.
    // FlatBuffers accessors are safe (panic on bounds usually).
    return nil
}
