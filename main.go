package main

import (
	"strconv"

	"github.com/fatih/structs"
)

const (
	rentabilität       = "/Users/christianhovenbitzer/Desktop/fremdkosten/rentabilität.xlsx"
	eingangsrechnungen = "/Users/christianhovenbitzer/Desktop/fremdkosten/eingangsrechnungen.xlsx"
	resultPath         = "/Users/christianhovenbitzer/Desktop/fremdkosten/result.xlsx"
)

var (
	rentColumns = []string{"A", "C", "G", "I", "L", "E"}
	inputExcel  Excel
	destExcel   Excel
)

func main() {
	destExcel = ExcelFile(resultPath, "result")
	inputExcel = ExcelFile(rentabilität, "")
	PrintHeader(&inputExcel, 0)

	data := FilterColumns(&inputExcel, rentColumns)
	projects := []Project{}
	for _, row := range data {
		projects = append(projects, Project{
			customer:                row[0],
			number:                  row[1],
			externalCostsChargeable: mustParse(row[2]),
			externalCosts:           mustParse(row[3]),
			income:                  mustParse(row[4]),
			revenue:                 mustParse(row[5]),
		})
	}

	for _, p := range projects {
		Add(&destExcel, &p)
	}

	destExcel.Save(resultPath)
}

// Project defines the necessary fields from "rentabilität"
type Project struct {
	customer                string
	number                  string
	externalCostsChargeable float32
	externalCosts           float32
	income                  float32
	revenue                 float32
}

// Columns returns the columnnames from struct Project
func (p *Project) Columns() []string {
	return structs.Names(p)
}

// Insert inserts values from struct Project
func (p *Project) Insert(excel *Excel) {
	pMap := structs.Map(p)
	nextRow := excel.NextRow()
	for _, n := range structs.Names(p) {
		coords := excel.CoordsForHeader(n)
		coords.row = nextRow
		excel.AddValue(coords, pMap[n])
	}
}

func mustParse(s string) float32 {
	v, err := strconv.ParseFloat(s, 32)
	if err != nil {
		panic("couldn't parse string")
	}
	return float32(v)
}
