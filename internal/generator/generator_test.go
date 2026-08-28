package generator

import (
	"path/filepath"
	"strings"
	"testing"

	"devis-api/internal/devis"

	"github.com/xuri/excelize/v2"
)

const testTemplatePath = "../../template/devis_template.xlsx"

// longValue is long enough to force every adaptive field to wrap onto
// several lines - well beyond anything a short real-world value would need,
// so it exercises the same wrapping/row-growth path a very long real name
// would trigger.
const longValue = "Jean-Baptiste-Alexandre-Maximilien-Bartholomew-Christophe-Wolfgang-VAN DER BERG-MONTGOMERY-DELACROIX-SAINT-EXUPERY-D'ARTAGNAN"

func baseTestDevis() devis.Devis {
	return devis.Devis{
		Civilite:       "M",
		Nom:            "Dupont",
		Prenom:         "Jean",
		DateNaissance:  "12/03/1950",
		DateDeces:      "28/06/2026",
		HeureDeces:     "14h30",
		VilleDeces:     "Nanterre",
		CodePostal:     "92000",
		DateAdmission:  "01/07/2026",
		HeureAdmission: "09h00",
		DateDepart:     "03/07/2026",
		HeureDepart:    "10h00",
		Operateur:      "Pompes Funèbres Martin",
		DateCommande:   "01/07/2026",
		Prestations:    map[string]float64{},
	}
}

// adaptiveFields lists every header cell whose font/wrapping is supposed to
// adapt to a long value, along with the row that must grow to fit it.
var adaptiveFields = []struct {
	name   string
	cell   string
	row    int
	mutate func(d *devis.Devis, value string)
}{
	{"Nom", "A6", 6, func(d *devis.Devis, v string) { d.Nom = v }},
	{"Prenom", "A10", 10, func(d *devis.Devis, v string) { d.Prenom = v }},
	{"NomNaissance", "A8", 8, func(d *devis.Devis, v string) { d.NomNaissance = v }},
	{"Operateur", "E7", 7, func(d *devis.Devis, v string) { d.Operateur = v }},
}

func generate(t *testing.T, d devis.Devis) *excelize.File {
	t.Helper()
	out := filepath.Join(t.TempDir(), "devis.xlsx")
	if err := GenerateXLSX(d, testTemplatePath, out); err != nil {
		t.Fatalf("GenerateXLSX: %v", err)
	}
	f, err := excelize.OpenFile(out)
	if err != nil {
		t.Fatalf("reopening generated file: %v", err)
	}
	t.Cleanup(func() { _ = f.Close() })
	return f
}

// TestGenerateXLSX_LongValuesGrowTheirRow checks that Nom, Prénom, Nom de
// naissance and Opérateur Funéraire each get a taller row when their value
// is long, instead of keeping the short-value row height and letting the
// text spill out of its bordered box.
func TestGenerateXLSX_LongValuesGrowTheirRow(t *testing.T) {
	for _, field := range adaptiveFields {
		t.Run(field.name, func(t *testing.T) {
			short := baseTestDevis()
			long := baseTestDevis()
			field.mutate(&long, longValue)

			shortFile := generate(t, short)
			longFile := generate(t, long)

			shortHeight, err := shortFile.GetRowHeight(SheetName, field.row)
			if err != nil {
				t.Fatalf("GetRowHeight (short): %v", err)
			}
			longHeight, err := longFile.GetRowHeight(SheetName, field.row)
			if err != nil {
				t.Fatalf("GetRowHeight (long): %v", err)
			}

			if longHeight <= shortHeight {
				t.Fatalf("expected row %d to grow for a long %s: short=%.1f long=%.1f", field.row, field.name, shortHeight, longHeight)
			}
		})
	}
}

// TestGenerateXLSX_LongValuesAreNotTruncated checks the full value survives
// in the cell (no silent truncation) for each adaptive field.
func TestGenerateXLSX_LongValuesAreNotTruncated(t *testing.T) {
	for _, field := range adaptiveFields {
		t.Run(field.name, func(t *testing.T) {
			d := baseTestDevis()
			field.mutate(&d, longValue)

			f := generate(t, d)
			got, err := f.GetCellValue(SheetName, field.cell)
			if err != nil {
				t.Fatalf("GetCellValue: %v", err)
			}
			if !strings.Contains(got, longValue) {
				t.Fatalf("expected %s (%s) to contain the full long value, got %q", field.name, field.cell, got)
			}
		})
	}
}

// TestGenerateXLSX_LongValuesWrap checks word-wrap is actually enabled on
// each adaptive field's cell - required for the extra row height to have any
// visible effect instead of leaving the overflowing text on one line.
func TestGenerateXLSX_LongValuesWrap(t *testing.T) {
	d := baseTestDevis()
	for _, field := range adaptiveFields {
		field.mutate(&d, longValue)
	}

	f := generate(t, d)
	for _, field := range adaptiveFields {
		t.Run(field.name, func(t *testing.T) {
			styleID, err := f.GetCellStyle(SheetName, field.cell)
			if err != nil {
				t.Fatalf("GetCellStyle: %v", err)
			}
			style, err := f.GetStyle(styleID)
			if err != nil {
				t.Fatalf("GetStyle: %v", err)
			}
			if style.Alignment == nil || !style.Alignment.WrapText {
				t.Fatalf("expected %s (%s) to have word-wrap enabled", field.name, field.cell)
			}
		})
	}
}

// TestGenerateXLSX_AllLongValuesAtOnceStillOnePage is the worst case: every
// adaptive field long at the same time. GenerateXLSX must still succeed and
// keep everything on the sheet's single print area (page count itself is a
// LibreOffice/PDF-export concern, out of scope for this pure-Go test, but a
// generation error here would mean the layout broke outright).
func TestGenerateXLSX_AllLongValuesAtOnceStillOnePage(t *testing.T) {
	d := baseTestDevis()
	for _, field := range adaptiveFields {
		field.mutate(&d, longValue)
	}
	_ = generate(t, d)
}

// TestGenerateXLSX_InitialesAreWrittenAtTheBottom checks the initials of the
// person who drafted the devis end up in the bottom box, and that the label
// alone is printed when no initials were entered so the box still reads as a
// field to fill in by hand.
func TestGenerateXLSX_InitialesAreWrittenAtTheBottom(t *testing.T) {
	for _, tc := range []struct {
		name      string
		initiales string
		want      string
	}{
		{"filled in", "JD", initialesLabel + " : JD"},
		{"left empty", "", initialesLabel},
	} {
		t.Run(tc.name, func(t *testing.T) {
			d := baseTestDevis()
			d.Initiales = tc.initiales

			f := generate(t, d)
			got, err := f.GetCellValue(SheetName, initialesCell)
			if err != nil {
				t.Fatalf("GetCellValue: %v", err)
			}
			if got != tc.want {
				t.Fatalf("expected %s to be %q, got %q", initialesCell, tc.want, got)
			}
		})
	}
}

// TestGenerateXLSX_InitialesBoxIsClosed checks the initials cell is drawn as
// a proper box: without a border on all four sides the line reads as loose
// text lost among the legal mentions, which is how it went unnoticed.
func TestGenerateXLSX_InitialesBoxIsClosed(t *testing.T) {
	d := baseTestDevis()
	d.Initiales = "JD"

	f := generate(t, d)
	styleID, err := f.GetCellStyle(SheetName, initialesCell)
	if err != nil {
		t.Fatalf("GetCellStyle: %v", err)
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		t.Fatalf("GetStyle: %v", err)
	}

	sides := map[string]bool{}
	for _, border := range style.Border {
		if border.Style > 0 {
			sides[border.Type] = true
		}
	}
	for _, side := range []string{"left", "top", "right", "bottom"} {
		if !sides[side] {
			t.Fatalf("expected %s to have a %s border, got %+v", initialesCell, side, style.Border)
		}
	}
}

// TestGenerateXLSX_InitialesAreInsidePrintArea guards the actual symptom:
// a value written outside the print area is silently dropped from the PDF
// export, so it never reaches the printed devis.
func TestGenerateXLSX_InitialesAreInsidePrintArea(t *testing.T) {
	d := baseTestDevis()
	d.Initiales = "JD"

	f := generate(t, d)
	names := f.GetDefinedName()
	var printArea string
	for _, name := range names {
		if name.Name == "_xlnm.Print_Area" && name.Scope == SheetName {
			printArea = name.RefersTo
		}
	}
	if printArea == "" {
		t.Fatalf("no print area defined for %s, got %+v", SheetName, names)
	}

	col, row, err := excelize.CellNameToCoordinates(initialesCell)
	if err != nil {
		t.Fatalf("CellNameToCoordinates: %v", err)
	}
	rangeRef := printArea[strings.Index(printArea, "!")+1:]
	firstCell, lastCell, ok := strings.Cut(strings.ReplaceAll(rangeRef, "$", ""), ":")
	if !ok {
		t.Fatalf("unexpected print area %q", printArea)
	}
	firstCol, firstRow, err := excelize.CellNameToCoordinates(firstCell)
	if err != nil {
		t.Fatalf("CellNameToCoordinates(%q): %v", firstCell, err)
	}
	lastCol, lastRow, err := excelize.CellNameToCoordinates(lastCell)
	if err != nil {
		t.Fatalf("CellNameToCoordinates(%q): %v", lastCell, err)
	}
	if col < firstCol || col > lastCol || row < firstRow || row > lastRow {
		t.Fatalf("expected %s to be inside print area %s", initialesCell, printArea)
	}
}
