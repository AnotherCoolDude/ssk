package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// Coordinates wraps coordinates in a struct
type Coordinates struct {
	row, column int
}

// ToString returns the coordinates as excelformatted string
func (c Coordinates) ToString() string {
	if c.row == 0 {
		c.row = 1
	}
	return fmt.Sprintf("%s%d", excelize.ToAlphaString(c.column), c.row)
}
