package main

import (
	"strconv"
	"strings"

	"github.com/AnotherCoolDude/excel"
	. "github.com/AnotherCoolDude/excel"
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
	destExcel = File(resultPath, "2018")
	rentExcel = File(rentabilität, "")
	erExcel = File(eingangsrechnungen, "")

	resultsMap = make(map[string]float32)
	smy = summary{}

	rentData := rentExcel.FirstSheet().FilterByColumn(rentColumns)
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

	erData := erExcel.FirstSheet().FilterByHeader(erHeader)

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
		destExcel.FirstSheet().Add(&p)
	}
	destExcel.FirstSheet().Add(&smy)

	destExcel.FirstSheet().FreezeHeader()

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

func (s *summary) Insert(sh *excel.Sheet) {

	tbStyle := excel.Style{Border: Top, Format: Euro}
	topBorderCell := Cell{Value: " ", Style: Style{Border: Top, Format: NoFormat}}

	totalCells := map[int]excel.Cell{
		0: Cell{Value: "Gesamt", Style: tbStyle},
		1: topBorderCell,
		2: Cell{Value: smy.tAR, Style: tbStyle},
		3: Cell{Value: smy.tWB, Style: tbStyle},
		4: Cell{Value: smy.tNWB, Style: tbStyle},
		5: Cell{Value: smy.tER, Style: tbStyle},
		6: topBorderCell,
		7: topBorderCell,
		8: Cell{Value: smy.tDB1, Style: tbStyle},
	}
	sh.AddEmptyRow()
	sh.AddRow(totalCells)
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
func (p *Project) Insert(sh *excel.Sheet) {

	currentPrefix := jobnrPrefix(p.number)
	lastPrefix := jobnrPrefix(lastProject.number)

	tbeStyle := Style{Border: Top, Format: Euro}
	topBorderCell := Cell{Value: " ", Style: Style{Border: Top, Format: NoFormat}}
	// check if current project is a new customer
	if currentPrefix != lastPrefix && lastPrefix != "" {
		sh.AddEmptyRow()
		tbnfStyle := Style{Border: Top, Format: NoFormat}

		customerSumCells := map[int]Cell{
			0: Cell{Value: lastProject.customer, Style: tbnfStyle},
			1: topBorderCell,
			2: Cell{Value: resultsMap["totalRevenues"], Style: tbeStyle},
			3: Cell{Value: resultsMap["totalExtCostChargeable"], Style: tbeStyle},
			4: Cell{Value: resultsMap["totalExtCost"], Style: tbeStyle},
			5: Cell{Value: resultsMap["totalER"], Style: tbeStyle},
			6: topBorderCell,
			7: topBorderCell,
			8: Cell{Value: resultsMap["db1"], Style: tbeStyle},
		}
		sh.AddRow(customerSumCells)
		sh.AddEmptyRow()
		sh.AddEmptyRow()

		smy.tAR += resultsMap["totalRevenues"]
		smy.tWB += resultsMap["totalExtCostChargeable"]
		smy.tNWB += resultsMap["totalExtCost"]
		smy.tER += resultsMap["totalER"]
		smy.tDB1 += resultsMap["db1"]
		resultsMap = make(map[string]float32)
	}

	projectCells := map[int]Cell{
		1: Cell{Value: p.number, Style: NoStyle()},
		2: Cell{Value: p.revenue, Style: EuroStyle()},
		3: Cell{Value: p.externalCostsChargeable, Style: EuroStyle()},
		4: Cell{Value: p.externalCosts, Style: EuroStyle()},
		8: Cell{Value: p.db1, Style: EuroStyle()},
	}
	sh.AddRow(projectCells)

	var sumER float32

	for i, fibu := range p.fibu {
		erCells := map[int]Cell{
			5: Cell{Value: p.invoice[i], Style: EuroStyle()},
			6: Cell{Value: fibu, Style: DateStyle()},
			7: Cell{Value: p.paginiernr[i], Style: IntegerStyle()},
		}
		sumER = sumER + p.invoice[i]
		sh.AddRow(erCells)
	}

	projectResultCells := map[int]Cell{
		2: Cell{Value: p.revenue, Style: tbeStyle},
		3: Cell{Value: p.externalCostsChargeable, Style: tbeStyle},
		4: Cell{Value: p.externalCosts, Style: tbeStyle},
		5: Cell{Value: sumER, Style: tbeStyle},
		6: topBorderCell,
		7: topBorderCell,
		8: Cell{Value: p.db1, Style: tbeStyle},
	}
	sh.AddRow(projectResultCells)
	sh.AddEmptyRow()

	resultsMap["totalRevenues"] += p.revenue
	resultsMap["totalExtCostChargeable"] += p.externalCostsChargeable
	resultsMap["totalExtCost"] += p.externalCosts
	resultsMap["totalER"] += sumER
	resultsMap["db1"] += p.db1

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
