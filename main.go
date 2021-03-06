package main

// TODO DB1 = AR - ER

import (
	"fmt"
	"strconv"
	"time"

	"github.com/AnotherCoolDude/excel"
	. "github.com/AnotherCoolDude/excel"
	"golang.org/x/text/language"
	"golang.org/x/text/message"
)

const (
	rentabilität         = "/Users/christianhovenbitzer/Desktop/fremdkosten/rent_18-feb19.xlsx"
	rentabilität19       = "/Users/christianhovenbitzer/Desktop/fremdkosten/rent_JanFeb19.xlsx"
	rentabilitätpr       = "/Users/christianhovenbitzer/Desktop/fremdkosten/rent_pr_18.xlsx"
	eingangsrechnungen   = "/Users/christianhovenbitzer/Desktop/fremdkosten/er_rechnungsbuch_17-19.xlsx"
	eingangsrechnungenpr = "/Users/christianhovenbitzer/Desktop/fremdkosten/er_rechnungsbuch_pr_17-19.xlsx"
	abgrenzung           = "/Users/christianhovenbitzer/Desktop/fremdkosten/01-2018.xlsx"
	einbeziehen          = "/Users/christianhovenbitzer/Desktop/fremdkosten/Abgrenzung Unfertige 2018.xlsx"
	adjust18             = "/Users/christianhovenbitzer/Desktop/fremdkosten/unfertige Leistungen 2017.xlsx"
	adjust19             = "/Users/christianhovenbitzer/Desktop/fremdkosten/unfertige Leistungen 2018.xlsx"

	resultPath = "/Users/christianhovenbitzer/Desktop/fremdkosten/result_18.xlsx"
)

var (
	rentColumns = []string{"A", "C", "E", "G", "I", "L", "E"}
	erColumns   = []string{"A", "F", "G", "K"}
	erHeader    = []string{"Paginiernummer", "Leistungsart", "FiBu Zeitraum", "Projekt Nr.", "Netto"}

	rentExcel   *Excel
	rent19Excel *Excel
	rentprExcel *Excel
	erExcel     *Excel
	erprExcel   *Excel
	destExcel   *Excel
	abgr18Excel *Excel
	abgr19Excel *Excel
	adj18Excel  *Excel
	adj19Excel  *Excel

	lastProject Project
	smy         summary
	customerSmy customerSummary

	printer = message.NewPrinter(language.Make("de"))

	adjustments = []adjustment{}
)

func main() {
	destExcel = File(resultPath, "2018")
	rentExcel = File(rentabilität, "")
	rent19Excel = File(rentabilität19, "")
	rentprExcel = File(rentabilitätpr, "")
	erExcel = File(eingangsrechnungen, "")
	erprExcel = File(eingangsrechnungenpr, "")
	abgr18Excel = File(abgrenzung, "")
	abgr19Excel = File(einbeziehen, "")
	adj18Excel = File(adjust18, "")
	adj19Excel = File(adjust19, "")

	smy = summary{}
	customerSmy = customerSummary{}

	// all data from rent
	rentData := rentExcel.FirstSheet().FilterByColumn(rentColumns)
	rentprData := rentprExcel.FirstSheet().FilterByColumn(rentColumns)

	// rent without SEIN Projects
	chargabeleProjects := [][]string{}
	for _, row := range rentData {
		if jobnrPrefix(row[1]) != "SEIN" {
			chargabeleProjects = append(chargabeleProjects, row)
		}
	}
	chargabeleProjectsPR := [][]string{}
	for _, row := range rentprData {
		if jobnrPrefix(row[1]) != "SEPR" {
			chargabeleProjectsPR = append(chargabeleProjectsPR, row)
		}
	}

	// rent cleaned from abgrenzung
	indicator := abgr18Excel.FirstSheet().FilterByHeader([]string{"Projektnummern", "Bemerkung"})
	reduced := []string{}
	for _, item := range indicator {
		if item[1] == "2017" {
			reduced = append(reduced, item[0])
		}
	}
	lastYearCleaned := [][]string{}
	for _, row := range chargabeleProjects {
		if hasIdenticalItem(row, reduced) {
			continue
		}
		lastYearCleaned = append(lastYearCleaned, row)
	}
	fmt.Printf("amount of items to remove: %d\n", len(reduced))

	//rent cleaned from 2019 projects, that have no connection to 2018
	rent19Data := rent19Excel.FirstSheet().FilterByColumn(rentColumns)

	indicator = abgr19Excel.FirstSheet().FilterByHeader([]string{"Projektnummern", "Jahr"})
	reduced = []string{}
	for _, item := range indicator {
		if item[1] == "2018" {
			reduced = append(reduced, item[0])
		}
	}

	for _, row := range rent19Data {
		if hasIdenticalItem(row, reduced) {
			lastYearCleaned = append(lastYearCleaned, row)
		}
	}
	fmt.Printf("amount of items to add: %d\n", len(reduced))

	fmt.Printf("amount of projcts to be added to the final file: %d\n", len(lastYearCleaned))

	// create objects
	projects := []Project{}
	for _, row := range lastYearCleaned {
		//fk := mustParseFloat(row[3]) + mustParseFloat(row[4])
		projects = append(projects, Project{
			customer:                row[0],
			number:                  row[1],
			externalCostsChargeable: mustParseFloat(row[3]),
			externalCosts:           mustParseFloat(row[4]),
			invoice:                 []float32{},
			fibu:                    []string{},
			paginiernr:              []string{},
			invoiceNr:               []string{},
			income:                  mustParseFloat(row[5]),
			revenue:                 mustParseFloat(row[2]),
			db1:                     mustParseFloat(row[2]),
		})
	}

	projectsPR := []Project{}
	for _, row := range chargabeleProjectsPR {
		//fk := mustParseFloat(row[3]) + mustParseFloat(row[4])
		projectsPR = append(projectsPR, Project{
			customer:                row[0],
			number:                  row[1],
			externalCostsChargeable: mustParseFloat(row[3]),
			externalCosts:           mustParseFloat(row[4]),
			invoice:                 []float32{},
			fibu:                    []string{},
			paginiernr:              []string{},
			invoiceNr:               []string{},
			income:                  mustParseFloat(row[5]),
			revenue:                 mustParseFloat(row[2]),
			db1:                     mustParseFloat(row[2]),
		})
	}

	// all data from er
	erData := erExcel.FirstSheet().FilterByHeader(erHeader)
	erprData := erprExcel.FirstSheet().FilterByHeader(erHeader)

	// include er into projects
	for _, row := range erData {
		for i, p := range projects {
			if row[3] == p.number {
				projects[i].paginiernr = append(projects[i].paginiernr, row[0])
				projects[i].invoiceNr = append(projects[i].invoiceNr, row[1])
				projects[i].fibu = append(projects[i].fibu, row[2])
				projects[i].invoice = append(projects[i].invoice, mustParseFloat(row[4]))
			}
		}
	}

	for _, row := range erprData {
		for i, p := range projectsPR {
			if row[3] == p.number {
				projectsPR[i].paginiernr = append(projectsPR[i].paginiernr, row[0])
				projectsPR[i].invoiceNr = append(projectsPR[i].invoiceNr, row[1])
				projectsPR[i].fibu = append(projectsPR[i].fibu, row[2])
				projectsPR[i].invoice = append(projectsPR[i].invoice, mustParseFloat(row[4]))
			}
		}
	}

	//adjust unfinished projects
	adjustments18 := adj18Excel.Sheet("konsolidiert").FilterByColumn([]string{"A", "B", "C"})
	adjustments19 := adj19Excel.Sheet("konsolidiert").FilterByColumn([]string{"A", "B", "C"})

	for _, row := range adjustments18 {
		adjustments = append(adjustments, adjustment{
			projectnr:     row[0],
			revenue:       mustParseFloat(row[1]) * -1,
			externalCosts: mustParseFloat(row[2]) * -1,
			note:          printer.Sprint("Anteil 2017"),
		})
	}

	for _, row := range adjustments19 {
		adjustments = append(adjustments, adjustment{
			projectnr:     row[0],
			revenue:       mustParseFloat(row[1]) * -1,
			externalCosts: mustParseFloat(row[2]),
			note:          printer.Sprint("Anteil 2019"),
		})
	}

	// since adjustments19 contains external costs from 18, we have to do this
	for _, p := range projects {
		for i, adj := range adjustments {
			if p.number == adj.projectnr {
				thisYear := adj.externalCosts
				adjustments[i].externalCosts = 0.0
				for _, er := range p.invoice {
					adjustments[i].externalCosts += er
				}
				adjustments[i].externalCosts -= thisYear
				adjustments[i].externalCosts = adjustments[i].externalCosts * -1
			}
		}
	}

	// insert data into new excel
	for _, p := range projects {
		currentPrefix := jobnrPrefix(p.number)
		lastPrefix := jobnrPrefix(lastProject.number)
		if currentPrefix != lastPrefix && lastPrefix != "" {
			destExcel.FirstSheet().Add(&customerSmy)
		}
		destExcel.FirstSheet().Add(&p)
	}
	destExcel.FirstSheet().Add(&customerSmy)
	destExcel.FirstSheet().Add(&smy)
	destExcel.FirstSheet().FreezeHeader()

	// reset globals
	customerSmy = customerSummary{}
	smy = summary{}
	lastProject = Project{}

	// insert PR data
	fmt.Printf("rows in projectsPR: %d\n", len(projectsPR))
	for _, p := range projectsPR {
		currentPrefix := jobnrPrefix(p.number)
		lastPrefix := jobnrPrefix(lastProject.number)
		if currentPrefix != lastPrefix && lastPrefix != "" {
			destExcel.Sheet("PR").Add(&customerSmy)
		}
		destExcel.Sheet("PR").Add(&p)
	}
	destExcel.Sheet("PR").Add(&customerSmy)
	destExcel.Sheet("PR").Add(&smy)
	destExcel.Sheet("PR").FreezeHeader()

	// save file
	destExcel.Save(resultPath)
}

// adjustment struct

type adjustment struct {
	projectnr     string
	revenue       float32
	externalCosts float32
	note          string
}

func (adj *adjustment) Columns() []string {
	return []string{}
}

func (adj *adjustment) Insert(sh *excel.Sheet) {
	adjCells := map[int]Cell{
		8: Cell{Value: adj.revenue, Style: EuroStyle()},
		9: Cell{Value: adj.externalCosts, Style: EuroStyle()},
		1: Cell{Value: adj.note, Style: NoStyle()},
	}
	sh.AddRow(adjCells)

}

// customer summary struct

type customerSummary struct {
	customer string
	ar       float32
	wb       float32
	nwb      float32
	er       float32
	abgrEL   float32
	abgrFK   float32
	db1      float32
	cellMap  map[string][]Coordinates
}

func (s *customerSummary) Columns() []string {
	return []string{}
}

func (s *customerSummary) Insert(sh *excel.Sheet) {
	tbnfStyle := Style{Border: Top, Format: NoFormat}
	tbCell := Cell{Value: " ", Style: tbnfStyle}
	tbeStyle := Style{Border: Top, Format: Euro}

	// Formulas
	arFormula := Formula{Coords: s.cellMap["ar"]}
	wbFormula := Formula{Coords: s.cellMap["wb"]}
	nwbFormula := Formula{Coords: s.cellMap["nwb"]}
	erFormula := Formula{Coords: s.cellMap["er"]}
	abgrELFormula := Formula{Coords: s.cellMap["abgrEL"]}
	abgrFKFormula := Formula{Coords: s.cellMap["abgrFK"]}
	db1Formula := Formula{Coords: s.cellMap["db1"]}

	customerSumCells := map[int]Cell{
		0:  Cell{Value: lastProject.customer, Style: tbnfStyle},
		1:  tbCell,
		2:  Cell{Value: arFormula.Add(), Style: tbeStyle},
		3:  Cell{Value: wbFormula.Add(), Style: tbeStyle},
		4:  Cell{Value: nwbFormula.Add(), Style: tbeStyle},
		5:  Cell{Value: erFormula.Add(), Style: tbeStyle},
		6:  tbCell,
		7:  tbCell,
		8:  Cell{Value: abgrELFormula.Add(), Style: tbeStyle},
		9:  Cell{Value: abgrFKFormula.Add(), Style: tbeStyle},
		10: Cell{Value: db1Formula.Add(), Style: tbeStyle},
	}
	sh.AddRow(customerSumCells)
	if smy.cellMap == nil {
		smy.cellMap = map[string][]Coordinates{}
	}
	smy.cellMap["ar"] = append(smy.cellMap["ar"], Coordinates{Row: sh.CurrentRow(), Column: 2})
	smy.cellMap["wb"] = append(smy.cellMap["wb"], Coordinates{Row: sh.CurrentRow(), Column: 3})
	smy.cellMap["nwb"] = append(smy.cellMap["nwb"], Coordinates{Row: sh.CurrentRow(), Column: 4})
	smy.cellMap["er"] = append(smy.cellMap["er"], Coordinates{Row: sh.CurrentRow(), Column: 5})
	smy.cellMap["abgrEL"] = append(smy.cellMap["abgrEL"], Coordinates{Row: sh.CurrentRow(), Column: 8})
	smy.cellMap["abgrFK"] = append(smy.cellMap["abgrFK"], Coordinates{Row: sh.CurrentRow(), Column: 9})
	smy.cellMap["db1"] = append(smy.cellMap["db1"], Coordinates{Row: sh.CurrentRow(), Column: 10})
	sh.AddEmptyRow()
	sh.AddEmptyRow()

	smy.tAR += s.ar
	smy.tWB += s.wb
	smy.tNWB += s.nwb
	smy.tER += s.er
	smy.tAbgrEL += s.abgrEL
	smy.tAbgrFK += s.abgrFK
	smy.tDB1 += s.db1

	customerSmy = customerSummary{}
	fmt.Printf("customer %s added\n", lastProject.customer)
}

// summary struct

type summary struct {
	tAR     float32
	tWB     float32
	tNWB    float32
	tER     float32
	tAbgrEL float32
	tAbgrFK float32
	tDB1    float32
	cellMap map[string][]Coordinates
}

func (s *summary) Columns() []string {
	return []string{}
}

func (s *summary) Insert(sh *excel.Sheet) {
	tbStyle := excel.Style{Border: Top, Format: Euro}
	topBorderCell := Cell{Value: " ", Style: Style{Border: Top, Format: NoFormat}}

	arFormula := Formula{Coords: s.cellMap["ar"]}
	wbFormula := Formula{Coords: s.cellMap["wb"]}
	nwbFormula := Formula{Coords: s.cellMap["nwb"]}
	erFormula := Formula{Coords: s.cellMap["er"]}
	abgrELFormula := Formula{Coords: s.cellMap["abgrEL"]}
	abgrFKFormula := Formula{Coords: s.cellMap["abgrFK"]}
	db1Formula := Formula{Coords: s.cellMap["db1"]}

	totalCells := map[int]excel.Cell{
		0:  Cell{Value: "Gesamt", Style: tbStyle},
		1:  topBorderCell,
		2:  Cell{Value: arFormula.Add(), Style: tbStyle},
		3:  Cell{Value: wbFormula.Add(), Style: tbStyle},
		4:  Cell{Value: nwbFormula.Add(), Style: tbStyle},
		5:  Cell{Value: erFormula.Add(), Style: tbStyle},
		6:  topBorderCell,
		7:  topBorderCell,
		8:  Cell{Value: abgrELFormula.Add(), Style: tbStyle},
		9:  Cell{Value: abgrFKFormula.Add(), Style: tbStyle},
		10: Cell{Value: db1Formula.Add(), Style: tbStyle},
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
	fibu                    []string
	paginiernr              []string
	invoiceNr               []string
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
		"Leistungsart",
		"Abgrenzung EL",
		"Abgrenzung FK",
		"DB 1",
	}
}

// 0=Kunde 1=Jobnr 2=AR Erlös 3=FKwb 4=FKnwb 5=Eingangsr 6=FiBu 7=Leistungsart 8=Abgr EL 9=Abgr FK 10=DB1

// Insert inserts values from struct Project
func (p *Project) Insert(sh *excel.Sheet) {
	tbeStyle := Style{Border: Top, Format: Euro}
	topBorderCell := Cell{Value: " ", Style: Style{Border: Top, Format: NoFormat}}

	projectCells := map[int]Cell{
		1: Cell{Value: p.number, Style: NoStyle()},
		2: Cell{Value: p.revenue, Style: EuroStyle()},
		3: Cell{Value: p.externalCostsChargeable, Style: EuroStyle()},
		4: Cell{Value: p.externalCosts, Style: EuroStyle()},
	}
	sh.AddRow(projectCells)

	var sumER float32

	invoiceCells := []Coordinates{}
	for i, fibu := range p.fibu {
		erCells := map[int]Cell{
			5: Cell{Value: p.invoice[i], Style: EuroStyle()},
			6: Cell{Value: fibu, Style: NoStyle()},
			7: Cell{Value: p.invoiceNr[i], Style: NoStyle()},
		}
		sumER = sumER + p.invoice[i]
		sh.AddRow(erCells)
		invoiceCells = append(invoiceCells, Coordinates{Row: sh.CurrentRow(), Column: 5})
	}

	p.db1 -= sumER

	adjRevCells := []Coordinates{}
	adjExtCostsCells := []Coordinates{}
	currentAdj := adjustment{}
	for _, adj := range adjustments {
		if p.number == adj.projectnr {
			adj.Insert(sh)
			currentAdj = adj
			// sumER += adj.externalCosts
			// p.revenue += adj.revenue
			// p.db1 = p.revenue - sumER
			adjRevCells = append(adjRevCells, Coordinates{Row: sh.CurrentRow(), Column: 8})
			adjExtCostsCells = append(adjExtCostsCells, Coordinates{Row: sh.CurrentRow(), Column: 9})
			reduction := adj.externalCosts + adj.revenue
			p.db1 += reduction
		}
	}

	invoiceFormula := Formula{Coords: invoiceCells}
	adjRevFormula := Formula{Coords: adjRevCells}
	adjExtCostsFormula := Formula{Coords: adjExtCostsCells}
	db1Formula := Formula{Coords: []Coordinates{
		Coordinates{Row: sh.NextRow(), Column: 2},
		Coordinates{Row: sh.NextRow(), Column: 5},
		Coordinates{Row: sh.NextRow(), Column: 8},
		Coordinates{Row: sh.NextRow(), Column: 9},
	}}
	projectResultCells := map[int]Cell{
		2: Cell{Value: p.revenue, Style: tbeStyle},
		3: Cell{Value: p.externalCostsChargeable, Style: tbeStyle},
		4: Cell{Value: p.externalCosts, Style: tbeStyle},
		5: Cell{Value: invoiceFormula.Add(), Style: tbeStyle},
		6: topBorderCell,
		7: topBorderCell,
		8: Cell{Value: adjRevFormula.Add(), Style: tbeStyle},
		9: Cell{Value: adjExtCostsFormula.Add(), Style: tbeStyle},
		10: Cell{Value: db1Formula.Raw(func(coords []Coordinates) string {
			return fmt.Sprintf("=%s-%s+%s+%s", coords[0].ToString(), coords[1].ToString(), coords[2].ToString(), coords[3].ToString())
		}), Style: tbeStyle},
	}
	sh.AddRow(projectResultCells)
	if customerSmy.cellMap == nil {
		customerSmy.cellMap = map[string][]Coordinates{}
	}
	customerSmy.cellMap["ar"] = append(customerSmy.cellMap["ar"], Coordinates{Row: sh.CurrentRow(), Column: 2})
	customerSmy.cellMap["wb"] = append(customerSmy.cellMap["wb"], Coordinates{Row: sh.CurrentRow(), Column: 3})
	customerSmy.cellMap["nwb"] = append(customerSmy.cellMap["nwb"], Coordinates{Row: sh.CurrentRow(), Column: 4})
	customerSmy.cellMap["er"] = append(customerSmy.cellMap["er"], Coordinates{Row: sh.CurrentRow(), Column: 5})
	customerSmy.cellMap["abgrEL"] = append(customerSmy.cellMap["abgrEL"], Coordinates{Row: sh.CurrentRow(), Column: 8})
	customerSmy.cellMap["abgrFK"] = append(customerSmy.cellMap["abgrFK"], Coordinates{Row: sh.CurrentRow(), Column: 9})
	customerSmy.cellMap["db1"] = append(customerSmy.cellMap["db1"], Coordinates{Row: sh.CurrentRow(), Column: 10})
	sh.AddEmptyRow()

	customerSmy.ar += p.revenue
	customerSmy.wb += p.externalCostsChargeable
	customerSmy.nwb += p.externalCosts
	customerSmy.er += sumER
	customerSmy.abgrEL += currentAdj.revenue
	customerSmy.abgrFK += currentAdj.externalCosts
	customerSmy.db1 += p.db1

	lastProject = *p
}

// removes the rows, that contain on if the indicator items. If reversed, removes all rows, that dont have an indicator
func removeRows(fromData [][]string, indicator []string, reversed bool) [][]string {
	fmt.Printf("amount rows input: %d\n", len(fromData))
	resultData := [][]string{}
	for _, row := range fromData {

		if !reversed {
			if hasIdenticalItem(row, indicator) {
				continue
			}
			resultData = append(resultData, row)
		} else {
			// needs some work
			if hasIdenticalItem(row, indicator) {
				resultData = append(resultData, row)
			}
		}

	}
	fmt.Printf("amount rows output: %d\n", len(resultData))
	return resultData
}

func hasIdenticalItem(slice, sliceToCompare []string) bool {
	for _, item := range slice {
		for _, compItem := range sliceToCompare {
			if item == compItem {
				return true
			}
		}
	}
	return false
}

func mustParseFloat(s string) float32 {
	v, err := strconv.ParseFloat(s, 32)
	if err != nil {
		fmt.Println(err)
		panic("couldn't parse string")
	}
	return float32(v)
}

func mustParseInt(s string) int {
	v, err := strconv.Atoi(s)
	if err != nil {
		fmt.Println(err)
		panic("couldn't parse string")
	}
	return v
}

func mustParseDate(s string) float32 {
	layout := "Jan-06"
	date, err := time.Parse(layout, s)
	if err != nil {
		fmt.Println(err)
		panic("couldn't parse date")
	}
	refDate, err := time.Parse("02-01-2006", "01-01-1900")
	if err != nil {
		fmt.Println(err)
		panic("couldn't parse date")
	}
	fmt.Printf("date: %s\nrefDate: %s\n", date.String(), refDate.String())
	duration := date.Sub(refDate)
	fmt.Printf("Duration: %s\n", duration.String())
	return float32(duration.Hours() / float64(24))
}

func jobnrPrefix(jobnr string) string {
	if jobnr == "" {
		return ""
	}
	return jobnr[:4]
}
