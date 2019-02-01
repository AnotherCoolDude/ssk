package main

import (
	"fmt"

	"github.com/buger/goterm"
)

// Table creates a Table from a provided 2D-Slice of Strings. The first inner Slice provides the Header
func Table(data [][]string) string {
	table := goterm.NewTable(0, 10, 5, ' ', 0)
	for _, row := range data {
		for _, item := range row {
			fmt.Fprintf(table, "%s\t", item)
		}
		fmt.Fprintf(table, "\n")
	}
	return table.String()
}
