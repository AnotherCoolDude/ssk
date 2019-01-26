package main

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// PrintHeader prints a table that contains the header of each sheet and it's index
func PrintHeader(path string, startingRow int) {
	file, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Printf("could not open file:\n%s\nerror: %s\n", path, err)
		return
	}
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
