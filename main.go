package main

import (
	"strconv"
	"strings"
)

const (
	rentabilität       = "/Users/christianhovenbitzer/Desktop/fremdkosten/rent_18.xlsx"
	eingangsrechnungen = "/Users/christianhovenbitzer/Desktop/fremdkosten/er_nov17-jan19.xlsx"
	resultPath         = "/Users/christianhovenbitzer/Desktop/fremdkosten/result_18.xlsx"
)

var (
	rentColumns = []string{"A", "C", "E", "G", "I", "L", "E"}
	erColumns   = []string{"A", "F", "G", "K"}
	erHeader    = []string{"Paginiernummer", "FiBu-Zeitraum", "Projektnummern", "Netto (Dokument)"}
	rentExcel   *Excel
	erExcel     *Excel
	destExcel   *Excel
	resultsMap  map[string]float32
	lastProject Project
	smy         summary
)

func main() {
	destExcel = ExcelFile(resultPath, "2018")
	rentExcel = ExcelFile(rentabilität, "")
	erExcel = ExcelFile(eingangsrechnungen, "")

	resultsMap = make(map[string]float32)
	smy = summary{}

	rentData := rentExcel.FilterByColumn(rentColumns)
	projects := []Project{}
	for _, row := range rentData {
		fk := mustParseFloat(row[3]) + mustParseFloat(row[4])
		projects = append(projects, Project{
			customer:                row[0],
			number:                  row[1],
			externalCostsChargeable: mustParseFloat(row[3]),
			externalCosts:           mustParseFloat(row[4]),
			invoice:                 []float32{},
			fibu:                    []int{},
			paginiernr:              []string{},
			income:                  mustParseFloat(row[5]),
			revenue:                 mustParseFloat(row[2]),
			db1:                     mustParseFloat(row[2]) - fk,
		})
	}

	erData := erExcel.FilterByHeader(erHeader)

	for _, row := range erData {

		for i, p := range projects {
			if row[2] == p.number {
				projects[i].paginiernr = append(projects[i].paginiernr, row[0])
				projects[i].fibu = append(projects[i].fibu, mustParseInt(row[1]))
				projects[i].invoice = append(projects[i].invoice, mustParseFloat(row[3]))
			}
		}
	}

	for _, p := range projects {
		destExcel.Add(&p)
	}
	destExcel.Add(&smy)

	destExcel.FreezeHeader()

	destExcel.Save(resultPath)
}

type summary struct {
	tAR  float32
	tWB  float32
	tNWB float32
	tER  float32
	tDB1 float32
}

func (s *summary) Columns() []string {
	return []string{}
}

func (s *summary) Insert(excel *Excel) {
	row := excel.NextRow() + 1

	excel.AddStyle([]Coordinates{
		Coordinates{column: 0, row: row},
		Coordinates{column: 8, row: row},
	}, BorderTop)

	excel.AddValue(Coordinates{column: 0, row: row}, "Gesamt", false)
	excel.AddValue(Coordinates{column: 2, row: row}, smy.tAR, false)
	excel.AddValue(Coordinates{column: 3, row: row}, smy.tWB, false)
	excel.AddValue(Coordinates{column: 4, row: row}, smy.tNWB, false)
	excel.AddValue(Coordinates{column: 5, row: row}, smy.tER, false)
	excel.AddValue(Coordinates{column: 8, row: row}, smy.tDB1, false)
	excel.AddCondition(Coordinates{column: 5, row: row}, smy.tNWB+smy.tWB)
}

// Project defines the necessary fields for the result xlsx
type Project struct {
	customer                string
	number                  string
	externalCostsChargeable float32
	externalCosts           float32
	invoice                 []float32
	fibu                    []int
	paginiernr              []string
	income                  float32
	revenue                 float32
	db1                     float32
}

// Columns returns the columnnames from struct Project
func (p *Project) Columns() []string {
	return []string{
		"Kunde",
		"Jobnr",
		"AR Erlös",
		"FK wb",
		"FK nwb",
		"ER Aufwendungen",
		"ER FiBu",
		"Paginiernr",
		"DB 1",
	}
}

// 0=Kunde 1=Jobnr 2=AR Erlös 3=FKwb 4=FKnwb 5=Eingangsr 6=FiBu 7=Pagnr 8=Umsatz vor El

// Insert inserts values from struct Project
func (p *Project) Insert(excel *Excel) {
	row := excel.NextRow()

	currentPrefix := jobnrPrefix(p.number)
	lastPrefix := jobnrPrefix(lastProject.number)

	// check if current project is a new customer
	if currentPrefix != lastPrefix && lastPrefix != "" {
		summaryRow := row + 1
		excel.AddStyle([]Coordinates{
			Coordinates{column: 0, row: summaryRow},
			Coordinates{column: 8, row: summaryRow},
		}, BorderTop)
		excel.AddValue(Coordinates{column: 0, row: summaryRow}, lastProject.customer, false)
		excel.AddValue(Coordinates{column: 2, row: summaryRow}, resultsMap["totalRevenues"], false)
		excel.AddValue(Coordinates{column: 3, row: summaryRow}, resultsMap["totalExtCostChargeable"], false)
		excel.AddValue(Coordinates{column: 4, row: summaryRow}, resultsMap["totalExtCost"], false)
		excel.AddValue(Coordinates{column: 5, row: summaryRow}, resultsMap["totalER"], false)
		excel.AddValue(Coordinates{column: 8, row: summaryRow}, resultsMap["db1"], false)
		excel.AddCondition(Coordinates{column: 5, row: summaryRow}, resultsMap["totalExtCostChargeable"]+resultsMap["totalExtCost"])
		excel.AddEmptyRow(summaryRow + 1)
		smy.tAR += resultsMap["totalRevenues"]
		smy.tWB += resultsMap["totalExtCostChargeable"]
		smy.tNWB += resultsMap["totalExtCost"]
		smy.tER += resultsMap["totalER"]
		smy.tDB1 += resultsMap["db1"]
		resultsMap = make(map[string]float32)
		row = summaryRow + 2
	}

	excel.AddValue(Coordinates{column: 1, row: row}, p.number, false)
	excel.AddValue(Coordinates{column: 2, row: row}, p.revenue, false)
	excel.AddValue(Coordinates{column: 3, row: row}, p.externalCostsChargeable, false)
	excel.AddValue(Coordinates{column: 4, row: row}, p.externalCosts, false)
	excel.AddValue(Coordinates{column: 8, row: row}, p.db1, false)

	var sumER float32

	for i, fibu := range p.fibu {
		sumER = sumER + p.invoice[i]
		excel.AddValue(Coordinates{column: 5, row: row + i + 1}, p.invoice[i], false)
		excel.AddValue(Coordinates{column: 6, row: row + i + 1}, fibu, true)
		excel.AddValue(Coordinates{column: 7, row: row + i + 1}, p.paginiernr[i], false)
	}

	resultRow := row + len(p.fibu) + 1

	excel.AddStyle([]Coordinates{
		Coordinates{column: 2, row: resultRow},
		Coordinates{column: 8, row: resultRow},
	}, BorderTop)

	excel.AddStyle([]Coordinates{
		Coordinates{column: 2, row: row},
		Coordinates{column: 2, row: resultRow - 1},
	}, BorderLeftRight)

	excel.AddStyle([]Coordinates{
		Coordinates{column: 4, row: row},
		Coordinates{column: 4, row: resultRow - 1},
	}, BorderRight)

	excel.AddStyle([]Coordinates{
		Coordinates{column: 7, row: row},
		Coordinates{column: 7, row: resultRow - 1},
	}, BorderRight)

	excel.AddValue(Coordinates{column: 2, row: resultRow}, p.revenue, false)
	resultsMap["totalRevenues"] += p.revenue
	excel.AddValue(Coordinates{column: 3, row: resultRow}, p.externalCostsChargeable, false)
	resultsMap["totalExtCostChargeable"] += p.externalCostsChargeable
	excel.AddValue(Coordinates{column: 4, row: resultRow}, p.externalCosts, false)
	resultsMap["totalExtCost"] += p.externalCosts
	excel.AddValue(Coordinates{column: 5, row: resultRow}, sumER, false)
	resultsMap["totalER"] += sumER
	excel.AddCondition(Coordinates{column: 5, row: resultRow}, p.externalCostsChargeable+p.externalCosts)
	excel.AddValue(Coordinates{column: 8, row: resultRow}, p.db1, false)
	resultsMap["db1"] += p.db1

	excel.AddEmptyRow(resultRow + 1)

	lastProject = *p
}

func mustParseFloat(s string) float32 {
	v, err := strconv.ParseFloat(s, 32)
	if err != nil {
		panic("couldn't parse string")
	}
	return float32(v)
}

func mustParseInt(s string) int {
	v, err := strconv.Atoi(s)
	if err != nil {
		panic("couldn't parse string")
	}
	return v
}

func jobnrPrefix(jobnr string) string {
	splited := strings.Split(jobnr, "-")
	return splited[0]
}
