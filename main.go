package main

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
	for i, row := range data {
		projects = append(projects, project{
			customer: row[0],
			number: row[1],
			
		})
	}
}

type project struct {
	customer                string
	number                  string
	externalCosts           float32
	externalCostsChargeable float32
	income                  float32
	revenue                 float32
}
