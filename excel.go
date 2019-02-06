package main

import (
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const (

	// NoBorder leaves the cell without border
	NoBorder BorderID = 0
	// Top adds a top border to the cell
	Top BorderID = 1
	// Left adds a left border to the cell
	Left BorderID = 2
	// Right adds a right border to the cell
	Right BorderID = 3
	// LeftRight adds a left and right border to the cell
	LeftRight BorderID = 4

	// NoFormat leaves the cell without format
	NoFormat FormatID = 0
	// Date formates the value of the cell to a date
	Date FormatID = 1
	// Euro formates the value of the cell to euro
	Euro FormatID = 2
	// Integer formates the value of the cell to integer
	Integer FormatID = 3
)

// Sheet wraps the sheets of a excel file into a struct
type Sheet struct {
	file *excelize.File
	name string
}

// Excel wraps the excelize package
type Excel struct {
	file            *excelize.File
	sheets          *[]Sheet
	ActiveSheetName string
}

// Sheet retruns the sheet by name or creates a new one
func (excel *Excel) Sheet(name string) *Sheet {
	// Sheet exists
	for _, existingSheet := range *excel.sheets {
		if existingSheet.name == name {
			return &existingSheet
		}
	}

	newSheet := Sheet{file: excel.file, name: name}
	excel.file.NewSheet(name)
	return &newSheet
}

// FirstSheet returns the first sheet found in the excel file
func (excel *Excel) FirstSheet() *Sheet {
	shs := excel.sheets
	return &(*shs)[0]
}

// ExcelFile opens/creates a Excel file. If newly created, names the first sheet after sheetname
func ExcelFile(path string, sheetname string) *Excel {
	var eFile *excelize.File
	var sheets []Sheet
	if _, err := os.Stat(path); os.IsNotExist(err) {
		fmt.Println("file not existing, creating new...")
		eFile = excelize.NewFile()
		sheetIndex := eFile.GetActiveSheetIndex()
		oldName := eFile.GetSheetName(sheetIndex)
		eFile.SetSheetName(oldName, sheetname)
		sheets = append(sheets, Sheet{file: eFile, name: sheetname})
	} else {
		eFile, err = excelize.OpenFile(path)
		sheetMap := eFile.GetSheetMap()
		for _, name := range sheetMap {
			sheets = append(sheets, Sheet{file: eFile, name: name})
		}
		if err != nil {
			fmt.Printf("couldn't open file at path\n%s\nerr: %s", path, err)
		}
	}
	return &Excel{
		file:            eFile,
		sheets:          &sheets,
		ActiveSheetName: eFile.GetSheetName(eFile.GetActiveSheetIndex()),
	}
}

// NextRow returns the next free Row
func (excel *Excel) NextRow() int {
	return len(excel.file.GetRows(excel.ActiveSheetName)) + 1
}

// Save saves the Excelfile to the provided path
func (excel *Excel) Save(path string) {
	excel.file.SaveAs(path)
}

// AddValue adds a value to the provided coordinates
func (excel *Excel) AddValue(coords Coordinates, value interface{}, style Style) {
	excel.file.SetCellValue(excel.ActiveSheetName, coords.ToString(), value)
	styleString := style.toString()
	if styleString == "" {
		return
	}
	st, err := excel.file.NewStyle(styleString)
	if err != nil {
		fmt.Println(styleString)
		fmt.Println(err)
	}
	excel.file.SetCellStyle(excel.ActiveSheetName, coords.ToString(), coords.ToString(), st)
}

// AddEmptyRow adds an empty row at index row
func (excel *Excel) AddEmptyRow() {
	freeRow := excel.NextRow()
	excel.file.SetCellStr(excel.ActiveSheetName, Coordinates{column: 0, row: freeRow}.ToString(), " ")
}

// AddCondition adds a condition, that fills the cell red if its value is less than comparison
func (excel *Excel) AddCondition(coord Coordinates, comparison float32) {
	compString := fmt.Sprintf("%f", comparison)
	format, err := excel.file.NewConditionalStyle(`{"fill":{"type":"pattern","color":["#F44E42"],"pattern":1}}`)
	if err != nil {
		fmt.Printf("couldn't create conditional style: %s\n", err)
	}
	excel.file.SetConditionalFormat(excel.ActiveSheetName, coord.ToString(), fmt.Sprintf(`[{"type":"cell","criteria":"<","format":%d,"value":%s}]`, format, compString))
}

// GetValue returns the Value from the cell at coord
func (excel *Excel) GetValue(coord Coordinates) string {
	return excel.file.GetCellValue(excel.ActiveSheetName, coord.ToString())
}

// FreezeHeader freezes the headerrow
func (excel *Excel) FreezeHeader() {
	excel.file.SetPanes(excel.ActiveSheetName, `{"freeze":true,"split":false,"x_split":0,"y_split":1,"top_left_cell":"A34","active_pane":"bottomLeft"}`)
}

// Style represents the style of a cell
type Style struct {
	Border BorderID
	Format FormatID
}

// BorderID represents the kind of Border
type BorderID int

// FormatID represents the formatting of the cell
type FormatID int

func (s Style) toString() string {
	st := ""

	if s.Border == NoBorder && s.Format == NoFormat {
		return st
	}

	switch s.Border {
	case NoBorder:
		st += `{`
	case Top:
		st += `{"border":[{"type":"top","color":"000000","style":1}]`
	case Left:
		st += `{"border":[{"type":"left","color":"000000","style":1}]`
	case Right:
		st += `{"border":[{"type":"right","color":"000000","style":1}]`
	case LeftRight:
		st += `{"border":[{"type":"left","color":"000000","style":1}, {"type":"right","color":"000000","style":1}]`
	}

	if s.Border != NoBorder && s.Format != NoFormat {
		st += `,`
	}

	switch s.Format {
	case NoFormat:
		st += `}`
	case Date:
		st += `"number_format": 17}`
	case Integer:
		st += `"number_format": 0}`
	case Euro:
		st += `"custom_number_format": "#,##0.00\\ [$\u20AC-1]"}`
	}
	return st
}

// DateStyle returns a Style struct that sets the cell to a date
func DateStyle() Style {
	return Style{
		Border: NoBorder,
		Format: Date,
	}
}

// EuroStyle returns a Style struct, that sets the cell to euro
func EuroStyle() Style {
	return Style{
		Border: NoBorder,
		Format: Euro,
	}
}

// NoStyle returns a Style struct, that doesn't modify the cell
func NoStyle() Style {
	return Style{
		Border: NoBorder,
		Format: NoFormat,
	}
}

// IntegerStyle returns a Style struct, that sets the cell to a integer
func IntegerStyle() Style {
	return Style{
		Border: NoBorder,
		Format: Integer,
	}
}

// Insertable defines Methods for structs to be insertable in a excelfile
type Insertable interface {
	Columns() []string
	Insert(sh *Sheet)
}

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

// PrintHeader prints a table that contains the header of each sheet and it's index
func PrintHeader(sh *Sheet, startingRow int) {
	if sh.isEmpty() {
		return
	}
	sheetMap := sh.file.GetSheetMap()
	for k, v := range sheetMap {
		headerTableData := [][]string{}
		headerTableData = append(headerTableData, []string{strconv.Itoa(k), v})
		rows := sh.file.GetRows(v)
		for index, head := range rows[startingRow] {
			headerTableData = append(headerTableData, []string{fmt.Sprintf("%s%d", excelize.ToAlphaString(index), startingRow+1), head})
		}
		t := Table(headerTableData)
		fmt.Print(t)
		fmt.Println()
	}

}

// FilterByHeader filters the excel file by its headertitle
func (sh *Sheet) FilterByHeader(header []string) [][]string {
	if sh.isEmpty() {
		return nil
	}

	data := sh.file.GetRows(sh.name)
	m := map[string]int{}

	for i, col := range data[0] {
		if contains(header, col) {
			m[col] = i
		}
	}
	sortedColumns := []string{}
	for _, h := range header {
		sortedColumns = append(sortedColumns, excelize.ToAlphaString(m[h]))
	}
	return sh.FilterByColumn(sortedColumns)
}

// FilterByColumn filters the excel file by its column
func (sh *Sheet) FilterByColumn(columns []string) [][]string {
	if sh.isEmpty() {
		return nil
	}
	data := sh.file.GetRows(sh.name)
	filteredData := [][]string{}

	for _, row := range data {
		filterMap := map[string]string{}
		for col, val := range row {
			if contains(columns, excelize.ToAlphaString(col)) {
				filterMap[excelize.ToAlphaString(col)] = val
			}
		}
		sortedRow := []string{}
		for _, c := range columns {
			sortedRow = append(sortedRow, filterMap[c])
		}
		filteredData = append(filteredData, sortedRow)
	}

	return filteredData[1:]
}

// Cell wraps a cell's value and it's style in a struct
type Cell struct {
	value interface{}
	style Style
}

// AddRow scanns for the next available row and inserts cells at the given indexes provided by the map
func (excel *Excel) AddRow(columnCellMap map[int]Cell) {
	freeRow := excel.NextRow()
	for col, cell := range columnCellMap {
		coords := Coordinates{column: col, row: freeRow}
		excel.file.SetCellValue(excel.ActiveSheetName, coords.ToString(), cell.value)
		styleString := cell.style.toString()
		if styleString == "" {
			continue
		}
		st, err := excel.file.NewStyle(styleString)
		if err != nil {
			fmt.Println(styleString)
			fmt.Println(err)
		}
		excel.file.SetCellStyle(excel.ActiveSheetName, coords.ToString(), coords.ToString(), st)
	}
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

func contains(slice []string, value string) bool {
	for _, v := range slice {
		if v == value {
			return true
		}
	}
	return false
}

func (sh *Sheet) isEmpty() bool {
	if len(sh.file.GetRows(sh.name)) == 0 {
		return true
	}
	return false
}
