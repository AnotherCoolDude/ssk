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

	tbStyle := Style{Border: Top, Format: Euro}
	topBorderCell := Cell{" ", Style{Border: Top, Format: NoFormat}}

	totalCells := map[int]Cell{
		0: Cell{"Gesamt", tbStyle},
		1: topBorderCell,
		2: Cell{smy.tAR, tbStyle},
		3: Cell{smy.tWB, tbStyle},
		4: Cell{smy.tNWB, tbStyle},
		5: Cell{smy.tER, tbStyle},
		6: topBorderCell,
		7: topBorderCell,
		8: Cell{smy.tDB1, tbStyle},
	}
	excel.AddEmpty()
	excel.AddRow(totalCells)
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
	// row := excel.NextRow()

	currentPrefix := jobnrPrefix(p.number)
	lastPrefix := jobnrPrefix(lastProject.number)

	tbeStyle := Style{Border: Top, Format: Euro}
	topBorderCell := Cell{" ", Style{Border: Top, Format: NoFormat}}
	// check if current project is a new customer
	if currentPrefix != lastPrefix && lastPrefix != "" {
		excel.AddEmpty()
		tbnfStyle := Style{Border: Top, Format: NoFormat}

		customerSumCells := map[int]Cell{
			0: Cell{lastProject.customer, tbnfStyle},
			1: topBorderCell,
			2: Cell{resultsMap["totalRevenues"], tbeStyle},
			3: Cell{resultsMap["totalExtCostChargeable"], tbeStyle},
			4: Cell{resultsMap["totalExtCost"], tbeStyle},
			5: Cell{resultsMap["totalER"], tbeStyle},
			6: topBorderCell,
			7: topBorderCell,
			8: Cell{resultsMap["db1"], tbeStyle},
		}
		excel.AddRow(customerSumCells)
		excel.AddEmpty()
		excel.AddEmpty()

		smy.tAR += resultsMap["totalRevenues"]
		smy.tWB += resultsMap["totalExtCostChargeable"]
		smy.tNWB += resultsMap["totalExtCost"]
		smy.tER += resultsMap["totalER"]
		smy.tDB1 += resultsMap["db1"]
		resultsMap = make(map[string]float32)
	}

	projectCells := map[int]Cell{
		1: Cell{p.number, NoStyle()},
		2: Cell{p.revenue, EuroStyle()},
		3: Cell{p.externalCostsChargeable, EuroStyle()},
		4: Cell{p.externalCosts, EuroStyle()},
		8: Cell{p.db1, EuroStyle()},
	}
	excel.AddRow(projectCells)

	var sumER float32

	for i, fibu := range p.fibu {
		erCells := map[int]Cell{
			5: Cell{p.invoice[i], EuroStyle()},
			6: Cell{fibu, DateStyle()},
			7: Cell{p.paginiernr[i], IntegerStyle()},
		}
		sumER = sumER + p.invoice[i]
		excel.AddRow(erCells)
	}

	projectResultCells := map[int]Cell{
		2: Cell{p.revenue, tbeStyle},
		3: Cell{p.externalCostsChargeable, tbeStyle},
		4: Cell{p.externalCosts, tbeStyle},
		5: Cell{sumER, tbeStyle},
		6: topBorderCell,
		7: topBorderCell,
		8: Cell{p.db1, tbeStyle},
	}
	excel.AddRow(projectResultCells)
	excel.AddEmpty()

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
