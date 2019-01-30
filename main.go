package main

import (
	"fmt"
	"strconv"

	"github.com/fatih/structs"
)

const (
	rentabilität       = "/Users/empfang/Desktop/fremdkosten/rentabilität.xlsx"
	eingangsrechnungen = "/Users/empfang/Desktop/fremdkosten/eingangsrechnungen.xlsx"
	resultPath         = "/Users/empfang/Desktop/fremdkosten/result.xlsx"
)

var (
	rentColumns = []string{"A", "C", "G", "I", "L", "E"}
	erColumns   = []string{"F", "G", "K"}
	rentExcel   Excel
	erExcel     Excel
	destExcel   Excel
)

func main() {
	destExcel = ExcelFile(resultPath, "result")
	rentExcel = ExcelFile(rentabilität, "")
	erExcel = ExcelFile(eingangsrechnungen, "")

	rentData := FilterColumns(&rentExcel, rentColumns)
	projects := []Project{}
	for _, row := range rentData {
		projects = append(projects, Project{
			customer:                row[0],
			number:                  row[1],
			externalCostsChargeable: mustParse(row[2]),
			externalCosts:           mustParse(row[3]),
			invoice:                 []float32{},
			fibu:                    []string{},
			income:                  mustParse(row[4]),
			revenue:                 mustParse(row[5]),
		})
	}

	erData := FilterColumns(&erExcel, erColumns)

	for _, row := range erData {

		for i, p := range projects {
			if row[1] == p.number {
				projects[i].fibu = append(projects[i].fibu, row[0])
				projects[i].invoice = append(projects[i].invoice, mustParse(row[2]))
			}
		}
	}

	for _, p := range projects {
		Add(&destExcel, &p)
	}

	destExcel.Save(resultPath)
}

// Project defines the necessary fields for the result xlsx
type Project struct {
	customer                string
	number                  string
	externalCostsChargeable float32
	externalCosts           float32
	invoice                 []float32
	fibu                    []string
	income                  float32
	revenue                 float32
}

// Columns returns the columnnames from struct Project
func (p *Project) Columns() []string {
	for _, n := range structs.Names(p) {
		fmt.Println(n)
	}
	return structs.Names(p)
}

// Insert inserts values from struct Project
func (p *Project) Insert(excel *Excel) {
	row := excel.NextRow()

	excel.AddValue(Coordinates{column: 0, row: row}, p.customer)
	excel.AddValue(Coordinates{column: 1, row: row}, p.number)
	excel.AddValue(Coordinates{column: 2, row: row}, p.externalCostsChargeable)
	excel.AddValue(Coordinates{column: 3, row: row}, p.externalCosts)
	excel.AddValue(Coordinates{column: 6, row: row}, p.income)
	excel.AddValue(Coordinates{column: 7, row: row}, p.revenue)

	for i, er := range p.fibu {
		excel.AddValue(Coordinates{column: 4, row: row + i + 1}, p.invoice[i])
		excel.AddValue(Coordinates{column: 5, row: row + i + 1}, er)
	}

}

func mustParse(s string) float32 {
	v, err := strconv.ParseFloat(s, 32)
	if err != nil {
		panic("couldn't parse string")
	}
	return float32(v)
}
