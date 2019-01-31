package main

import (
	"strconv"
	"strings"
)

const (
	rentabilität       = "/Users/christianhovenbitzer/Desktop/fremdkosten/rentabilität.xlsx"
	eingangsrechnungen = "/Users/christianhovenbitzer/Desktop/fremdkosten/eingangsrechnungen.xlsx"
	resultPath         = "/Users/christianhovenbitzer/Desktop/fremdkosten/result.xlsx"
)

var (
	rentColumns = []string{"A", "C", "E", "G", "I", "L", "E"}
	erColumns   = []string{"A", "F", "G", "K"}
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
			externalCostsChargeable: mustParse(row[3]),
			externalCosts:           mustParse(row[4]),
			invoice:                 []float32{},
			fibu:                    []string{},
			paginiernr:              []string{},
			income:                  mustParse(row[5]),
			revenue:                 mustParse(row[2]),
			revBevorOwnPerf:         mustParse(row[3]) - mustParse(row[4]),
		})
	}

	erData := FilterColumns(&erExcel, erColumns)

	for _, row := range erData {

		for i, p := range projects {
			if row[2] == p.number {
				projects[i].paginiernr = append(projects[i].paginiernr, row[0])
				projects[i].fibu = append(projects[i].fibu, row[1])
				projects[i].invoice = append(projects[i].invoice, mustParse(row[3]))
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
	paginiernr              []string
	income                  float32
	revenue                 float32
	revBevorOwnPerf         float32
}

// Columns returns the columnnames from struct Project
func (p *Project) Columns() []string {
	// for _, n := range structs.Names(p) {
	// 	fmt.Println(n)
	// }
	// return structs.Names(p)
	return []string{
		"Kunde",
		"Jobnr",
		"AR Erlös",
		"FK wb",
		"FK nwb",
		"ER Aufwendungen",
		"ER FiBu",
		"Paginiernr",
		"Umsatz vor EL",
	}
}

// 0=Kunde 1=Jobnr 2=Erlös 3=FKwb 4=FKnwb 5=Eingangsr 6=FiBu 7=Pagnr 8=Umsatz

// Insert inserts values from struct Project
func (p *Project) Insert(excel *Excel) {
	var newCustomer bool
	row := excel.NextRow()

	if row-2 > 0 {
		currentPrefix := jobnrPrefix(p.number)
		lastPrefix := jobnrPrefix(excel.GetValue(Coordinates{column: 1, row: row - 2}))
		if currentPrefix != lastPrefix {
			newCustomer = true
		} else {
			newCustomer = false
		}
	}

	excel.AddValue(Coordinates{column: 1, row: row}, p.number)
	excel.AddValue(Coordinates{column: 2, row: row}, p.revenue)
	excel.AddValue(Coordinates{column: 3, row: row}, p.externalCostsChargeable)
	excel.AddValue(Coordinates{column: 4, row: row}, p.externalCosts)
	excel.AddValue(Coordinates{column: 8, row: row}, p.revBevorOwnPerf)

	var sumER float32

	for i, er := range p.fibu {
		sumER = sumER + p.invoice[i]
		excel.AddValue(Coordinates{column: 5, row: row + i + 1}, p.invoice[i])
		excel.AddValue(Coordinates{column: 6, row: row + i + 1}, er)
		excel.AddValue(Coordinates{column: 7, row: row + i + 1}, p.paginiernr[i])
	}

	resultRow := row + len(p.fibu) + 2

	excel.AddStyle([]Coordinates{
		Coordinates{column: 2, row: resultRow},
		Coordinates{column: 8, row: resultRow},
	}, BorderTop)

	excel.AddStyle([]Coordinates{
		Coordinates{column: 2, row: row},
		Coordinates{column: 2, row: resultRow - 1},
	}, BorderLeftRight)

	excel.AddValue(Coordinates{column: 2, row: resultRow}, p.revenue)
	excel.AddValue(Coordinates{column: 3, row: resultRow}, p.externalCostsChargeable)
	excel.AddValue(Coordinates{column: 4, row: resultRow}, p.externalCosts)
	excel.AddValue(Coordinates{column: 5, row: resultRow}, sumER)
	excel.AddValue(Coordinates{column: 8, row: resultRow}, p.revBevorOwnPerf)

	excel.AddEmptyRow(resultRow + 1)

	if newCustomer {
		summaryRow := resultRow + 2
		excel.AddStyle([]Coordinates{
			Coordinates{column: 0, row: summaryRow},
			Coordinates{column: 8, row: summaryRow},
		}, BorderTop)
		excel.AddValue(Coordinates{column: 0, row: summaryRow}, p.customer)
		excel.AddValue(Coordinates{column: 0, row: summaryRow}, p.customer)
	}

}

func mustParse(s string) float32 {
	v, err := strconv.ParseFloat(s, 32)
	if err != nil {
		panic("couldn't parse string")
	}
	return float32(v)
}

func jobnrPrefix(jobnr string) string {
	splited := strings.Split(jobnr, "-")
	return splited[0]
}
