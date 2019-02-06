package main

import "fmt"

// Insertable defines Methods for structs to be insertable in a excelfile
type Insertable interface {
	Columns() []string
	Insert(sh *Sheet)
}

// Add inserts a insertable struct into a given file.
func (sh *Sheet) Add(data Insertable) {
	if sh.isEmpty() {
		fmt.Println("file is empty, adding header")
		headerCoords := Coordinates{row: 0, column: 0}
		for _, col := range data.Columns() {
			fmt.Printf("writing header %s at %s\n", col, headerCoords.ToString())
			sh.file.SetCellStr(sh.name, headerCoords.ToString(), col)
			headerCoords.column = headerCoords.column + 1
		}
	}
	data.Insert(sh)
}
