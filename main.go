package main

import (
	"strconv"

	"github.com/fatih/structs"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const (
	rentabilität       = "/Users/christianhovenbitzer/Desktop/fremdkosten/rentabilität.xlsx"
	eingangsrechnungen = "/Users/christianhovenbitzer/Desktop/fremdkosten/eingangsrechnungen.xlsx"
)

var (
	rentColumns = []string{"A1", "C1", "G1", "I1", "L1", "E1"}
)

func main() {
	PrintHeader(rentabilität, 0)
	PrintHeader(eingangsrechnungen, 0)

	data := FilterColumns(rentabilität, ActiveSheetname(rentabilität), rentColumns)
	projects := []project{}
	for _, row := range data {
		projects = append(projects, project{
			customer:                row[0],
			number:                  row[1],
			externalCostsChargeable: mustParse(row[2]),
			externalCosts:           mustParse(row[3]),
			income:                  mustParse(row[4]),
			revenue:                 mustParse(row[5]),
		})
	}

}

type insertable interface {
	insert(row int, file *excelize.File)
}

type project struct {
	customer                string
	number                  string
	externalCostsChargeable float32
	externalCosts           float32
	income                  float32
	revenue                 float32
}

func (p *project) insert(row int, file *excelize.File) {
	names := structs.Names(p)
	pMap := structs.Map(p)
	for i, name := range names {
		file.SetCellValue(ActiveSheetname(file.Path), Coords(i, row), pMap[name])
	}

}

func mustParse(s string) float32 {
	v, err := strconv.ParseFloat(s, 32)
	if err != nil {
		panic("couldn't parse string")
	}
	return float32(v)
}
