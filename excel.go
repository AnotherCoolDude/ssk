package main

import (
	"fmt"
	"os"
	"regexp"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// Excel wraps the excelize package
type Excel struct {
	File            *excelize.File
	ActiveSheetName string
}

// NextRow returns the next free Row
func (excel *Excel) NextRow() int {
	rows := excel.File.GetRows(excel.ActiveSheetName)
	return len(rows)
}

// Save saves the Excelfile to the provided path
func (excel *Excel) Save(path string) {
	for _, row := range excel.File.GetRows(excel.ActiveSheetName) {
		fmt.Printf("saving row: %s\n", row)
	}
	excel.File.SaveAs(path)
}

// CoordsForHeader returns the coords for the next free row and the given header
func (excel *Excel) CoordsForHeader(header string) Coordinates {
	rows := excel.File.GetRows(excel.ActiveSheetName)
	var freeRow int
	headerColumn := -1
	for index, row := range rows {
		//fmt.Println(row)
		if emptySlice(row) && headerColumn == -1 {
			fmt.Println("row is empty")
			continue
		}
		for j, column := range row {
			if column == header && headerColumn == -1 {
				headerColumn = j
				continue
			}
		}
		if headerColumn != -1 {
			if row[headerColumn] == "" {
				freeRow = index
				coords := Coordinates{
					column: headerColumn,
					row:    freeRow + 1,
				}
				fmt.Printf("next free coords for header %s: %s\n", header, coords.CoordString())
			}
		}

	}
	//fmt.Println("couldn't determine next free Cell")
	return Coordinates{
		column: 0,
		row:    0,
	}
}

// CoordsForHeader returns the coords for the next free row and the given header
// func (excel *Excel) CoordsForHeader(header string) Coordinates {
// 	rows := excel.File.GetRows(excel.ActiveSheetName)
// 	for i, column := range rows[0] {
// 		//fmt.Println(column)
// 		if column == header {
// 			freeRowIndex := 1
// 			for j := range rows {
// 				if excel.File.GetCellValue(excel.ActiveSheetName, Coordinates{column: i, row: freeRowIndex}.CoordString()) == "" {
// 					freeRowIndex = freeRowIndex + j
// 				}
// 			}
// 			coords := Coordinates{column: i, row: freeRowIndex}
// 			fmt.Printf("next free cell for header %s is at %s\n", header, coords.CoordString())
// 			return coords
// 		}
// 	}
// 	return Coordinates{0, 0}
// }

// AddValue adds a value to the provided coordinates
func (excel *Excel) AddValue(coords Coordinates, value interface{}) {
	excel.File.SetCellValue(excel.ActiveSheetName, coords.CoordString(), value)
	//fmt.Printf("Cell Value: %s\n", excel.File.GetCellValue(excel.ActiveSheetName, coords.CoordString()))
}

// ExcelFile opens/creates a Excel File. If newly created, names the first sheet after sheetname
func ExcelFile(path string, sheetname string) Excel {
	if _, err := os.Stat(path); os.IsNotExist(err) {
		fmt.Println("file not existing, creating new...")
		eFile := excelize.NewFile()
		sheetIndex := eFile.GetActiveSheetIndex()
		oldName := eFile.GetSheetName(sheetIndex)
		eFile.SetSheetName(oldName, "result")
		return Excel{
			File:            eFile,
			ActiveSheetName: eFile.GetSheetName(eFile.GetActiveSheetIndex()),
		}
	}
	eFile, _ := excelize.OpenFile(path)
	return Excel{
		File:            eFile,
		ActiveSheetName: eFile.GetSheetName(eFile.GetActiveSheetIndex()),
	}
}

// Insertable defines Methods for structs to be insertable in a excelfile
type Insertable interface {
	Columns() []string
	Insert(excel *Excel)
}

// Coordinates wraps coordinates in a struct
type Coordinates struct {
	row, column int
}

// CoordsTransmutable provides methods to better work with Coordinates in excel
type CoordsTransmutable interface {
	CoordString() string
	ColumnString() string
	ExistsIn(coords []Coordinates) bool
	ColumnExistsIn(coords []Coordinates) bool
}

// CoordString returns the coordinates as excelformatted string
func (c Coordinates) CoordString() string {
	c.row = c.row + 1
	return fmt.Sprintf("%s%d", excelize.ToAlphaString(c.column), c.row)
}

// ColumnString returns the column as excelformatted string
func (c Coordinates) ColumnString() string {
	return fmt.Sprintf("%s", excelize.ToAlphaString(c.column))
}

// ExistsIn checks, if coords contains c
func (c Coordinates) ExistsIn(coords []Coordinates) bool {
	for _, coord := range coords {
		if c.CoordString() == coord.CoordString() {
			return true
		}
	}
	return false
}

// ColumnExistsIn checks, if c's column exists in coords
func (c Coordinates) ColumnExistsIn(coords []Coordinates) bool {
	for _, coord := range coords {
		if c.ColumnString() == coord.ColumnString() {
			return true
		}
	}
	return false
}

// Coordinates returns a Coordinates Struct from a string
func coordinates(s string) Coordinates {
	reg := regexp.MustCompile("[0-9]+|[A-Z]+")
	result := reg.FindAllString(s, 2)
	n, _ := strconv.Atoi(result[1])
	return Coordinates{
		row:    n,
		column: excelize.TitleToNumber(result[0]),
	}
}

// PrintHeader prints a table that contains the header of each sheet and it's index
func PrintHeader(excel *Excel, startingRow int) {
	if excel.isEmpty() {
		return
	}
	sheetMap := excel.File.GetSheetMap()
	for k, v := range sheetMap {
		headerTableData := [][]string{}
		headerTableData = append(headerTableData, []string{strconv.Itoa(k), v})
		rows := excel.File.GetRows(v)
		for index, head := range rows[startingRow] {
			headerTableData = append(headerTableData, []string{fmt.Sprintf("%s%d", excelize.ToAlphaString(index), startingRow+1), head})
		}
		t := Table(headerTableData)
		fmt.Print(t)
		fmt.Println()
	}

}

// FilterColumns returns the values from all rows and selected columns
func FilterColumns(excel *Excel, columns []string) [][]string {
	if excel.isEmpty() {
		return nil
	}
	data := excel.File.GetRows(excel.ActiveSheetName)
	filteredData := [][]string{}
	filterCoords := []Coordinates{}
	for _, col := range columns {
		filterCoords = append(filterCoords, Coordinates{row: 0, column: excelize.TitleToNumber(col)})
	}

	for i, row := range data {
		filteredRow := []string{}
		for j, col := range row {
			currentCoords := Coordinates{row: i, column: j}
			if currentCoords.ColumnExistsIn(filterCoords) {
				filteredRow = append(filteredRow, col)
			}
		}
		filteredData = append(filteredData, filteredRow)
	}
	return filteredData[1:]
}

// Coords return the coord string for the provided column and row
func Coords(col, row int) string {
	alpha := excelize.ToAlphaString(col)
	return fmt.Sprintf("%s%d", alpha, row)
}

// Add inserts a insertable struct into a given file.
func Add(excel *Excel, data Insertable) {
	//rows := excel.File.GetRows(excel.ActiveSheetName)
	if excel.isEmpty() {
		fmt.Println("file is empty, adding header")
		//fmt.Println(excel.File.GetCellValue(excel.ActiveSheetName, "A1"))
		headerCoords := Coordinates{row: 0, column: 0}
		for _, col := range data.Columns() {
			fmt.Printf("writing header %s at %s\n", col, headerCoords.CoordString())
			excel.File.SetCellStr(excel.ActiveSheetName, headerCoords.CoordString(), col)
			headerCoords.column = headerCoords.column + 1
		}
		fmt.Println(excel.File.GetRows(excel.ActiveSheetName)[0])
	}
	data.Insert(excel)
}

func contains(slice []string, value string) bool {
	for _, v := range slice {
		if v == value {
			return true
		}
	}
	return false
}

func (excel *Excel) isEmpty() bool {
	if len(excel.File.GetRows(excel.ActiveSheetName)) == 0 {
		fmt.Println("file is empty")
		return true
	}
	return false
}

func emptySlice(slice []string) bool {
	for _, s := range slice {
		if s != "" {
			return false
		}
	}
	return true
}
