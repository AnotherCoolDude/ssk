package main

// Constants

const (

	// NoBorder leaves the cell without border
	NoBorder BorderID = 0
	// Top adds a top border to the cell
	Top BorderID = 1
	// Left adds a left border to the cell
	Left BorderID = 2
	// Right adds a right border to the cell
	Right BorderID = 3
	// LeftRight adds a left and right border to the cell
	LeftRight BorderID = 4

	// NoFormat leaves the cell without format
	NoFormat FormatID = 0
	// Date formates the value of the cell to a date
	Date FormatID = 1
	// Euro formates the value of the cell to euro
	Euro FormatID = 2
	// Integer formates the value of the cell to integer
	Integer FormatID = 3
)

// Structs

// Style represents the style of a cell
type Style struct {
	Border BorderID
	Format FormatID
}

// BorderID represents the kind of Border
type BorderID int

// FormatID represents the formatting of the cell
type FormatID int

// encode structs to string

func (s Style) toString() string {
	st := ""

	if s.Border == NoBorder && s.Format == NoFormat {
		return st
	}

	switch s.Border {
	case NoBorder:
		st += `{`
	case Top:
		st += `{"border":[{"type":"top","color":"000000","style":1}]`
	case Left:
		st += `{"border":[{"type":"left","color":"000000","style":1}]`
	case Right:
		st += `{"border":[{"type":"right","color":"000000","style":1}]`
	case LeftRight:
		st += `{"border":[{"type":"left","color":"000000","style":1}, {"type":"right","color":"000000","style":1}]`
	}

	if s.Border != NoBorder && s.Format != NoFormat {
		st += `,`
	}

	switch s.Format {
	case NoFormat:
		st += `}`
	case Date:
		st += `"number_format": 17}`
	case Integer:
		st += `"number_format": 0}`
	case Euro:
		st += `"custom_number_format": "#,##0.00\\ [$\u20AC-1]"}`
	}
	return st
}

// Convenience

// DateStyle returns a Style struct that sets the cell to a date
func DateStyle() Style {
	return Style{
		Border: NoBorder,
		Format: Date,
	}
}

// EuroStyle returns a Style struct, that sets the cell to euro
func EuroStyle() Style {
	return Style{
		Border: NoBorder,
		Format: Euro,
	}
}

// NoStyle returns a Style struct, that doesn't modify the cell
func NoStyle() Style {
	return Style{
		Border: NoBorder,
		Format: NoFormat,
	}
}

// IntegerStyle returns a Style struct, that sets the cell to a integer
func IntegerStyle() Style {
	return Style{
		Border: NoBorder,
		Format: Integer,
	}
}
