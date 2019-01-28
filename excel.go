package main

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// PrintHeader prints a table that contains the header of each sheet and it's index
func PrintHeader(path string, startingRow int) {
	file := open(path)
	sheetMap := file.GetSheetMap()

	for k, v := range sheetMap {
		headerTableData := [][]string{}
		headerTableData = append(headerTableData, []string{strconv.Itoa(k), v})
		rows := file.GetRows(v)
		for index, head := range rows[startingRow] {
			headerTableData = append(headerTableData, []string{fmt.Sprintf("%s%d", excelize.ToAlphaString(index), startingRow+1), head})
		}
		t := Table(headerTableData)
		fmt.Print(t)
		fmt.Println()
	}

}

// FilterColumns returns the values from all rows and selected columns
func FilterColumns(path, sheet string, columns []string) [][]string {
	file := open(path)
	data := file.GetRows(sheet)
	filteredData := [][]string{}
	for _, row := range data {
		filteredRow := []string{}
		for i, col := range row {
			if contains(columns, excelize.ToAlphaString(i)) {
				filteredRow = append(filteredRow, col)
			}
		}
		filteredData = append(filteredData, filteredRow)
	}
	return filteredData[1:]
}

// ActiveSheetname returns the sheetname of the active sheet
func ActiveSheetname(path string) string {
	file := open(path)
	names := file.GetSheetMap()
	return names[file.GetActiveSheetIndex()]
}

// Coords return the coord string for the provided column and row
func Coords(col, row int) string {
	alpha := excelize.ToAlphaString(col)
	return fmt.Sprintf("%s%d", alpha, row)
}

func open(path string) *excelize.File {
	file, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Printf("could not open file:\n%s\nerror: %s\n", path, err)
		return nil
	}
	return file
}

func contains(slice []string, value string) bool {
	for _, v := range slice {
		if v == value {
			return true
		}
	}
	return false
}
