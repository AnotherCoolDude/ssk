package main

import (
	"fmt"
	"os"
	"regexp"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const (
	// Sum is used to sum up the value of cells
	Sum Method = 0

	// BorderTop adds a border to the top of the cell
	BorderTop StyleType = 0
)

// Excel wraps the excelize package
type Excel struct {
	File            *excelize.File
	ActiveSheetName string
}

// Formula wraps a formula into a struct
type Formula struct {
	CoordsRange []Coordinates
	Method      Method
}

// Method represents the methods, than can be performed by a formula
type Method int

func (m Method) toString(coords []Coordinates) string {
	switch m {
	case Sum:
		return fmt.Sprintf("=SUMME(%s:%s)", coords[0].CoordString(), coords[1].CoordString())
	default:
		fmt.Println("unknown Method used...")
		return ""
	}
}

// StyleType defines the types a cell can be styled with
type StyleType int

func (st StyleType) toString() string {
	switch st {
	case BorderTop:
		return fmt.Sprintf(`{"border":[{"type":"top","color":"000000","style":1}]}`)
	default:
		fmt.Println("unknown Style used...")
		return ""
	}
}

// NextRow returns the next free Row
func (excel *Excel) NextRow() int {
	return len(excel.File.GetRows(excel.ActiveSheetName)) + 1
}

// Save saves the Excelfile to the provided path
func (excel *Excel) Save(path string) {
	excel.File.SaveAs(path)
}

// ColumnForHeader returns the column for the proived header
func (excel *Excel) ColumnForHeader(header string) string {
	headerCol := excel.File.GetRows(excel.ActiveSheetName)[0]
	for i, head := range headerCol {
		if head == header {
			return excelize.ToAlphaString(i)
		}
	}
	fmt.Printf("couldn't find header %s\n", header)
	return ""
}

// AddValue adds a value to the provided coordinates
func (excel *Excel) AddValue(coords Coordinates, value interface{}) {
	excel.File.SetCellValue(excel.ActiveSheetName, coords.CoordString(), value)
}

// AddFormula adds a formula to the provided coordinates
func (excel *Excel) AddFormula(coords Coordinates, formula Formula) {
	if coords.CoordString() == formula.CoordsRange[1].CoordString() || coords.CoordString() == formula.CoordsRange[0].CoordString() {
		v := excel.File.GetCellValue(excel.ActiveSheetName, formula.CoordsRange[0].CoordString())
		excel.File.SetCellValue(excel.ActiveSheetName, coords.CoordString(), v)
		return
	}
	excel.File.SetCellFormula(excel.ActiveSheetName, coords.CoordString(), formula.Method.toString(formula.CoordsRange))
}

// AddStyle adds a Style to the range of the provided coordinates
func (excel *Excel) AddStyle(coordsRange []Coordinates, styleType StyleType) {
	style, err := excel.File.NewStyle(styleType.toString())
	if err != nil {
		fmt.Println(err)
	}
	excel.File.SetCellStyle(excel.ActiveSheetName, coordsRange[0].CoordString(), coordsRange[1].CoordString(), style)
}

// AddEmptyRow adds an empty row at index row
func (excel *Excel) AddEmptyRow(row int) {
	excel.File.SetCellStr(excel.ActiveSheetName, Coordinates{column: 0, row: row}.CoordString(), " ")
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
	if c.row == 0 {
		c.row = 1
	}
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

// CoordsFromString returns a Coordinates Struct from a string
func CoordsFromString(s string) Coordinates {
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
	if excel.isEmpty() {
		fmt.Println("file is empty, adding header")
		headerCoords := Coordinates{row: 0, column: 0}
		for _, col := range data.Columns() {
			fmt.Printf("writing header %s at %s\n", col, headerCoords.CoordString())
			excel.File.SetCellStr(excel.ActiveSheetName, headerCoords.CoordString(), col)
			headerCoords.column = headerCoords.column + 1
		}
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
