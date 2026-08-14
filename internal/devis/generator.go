package devis

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Name of the sheet used in the template workbook.
const SheetName = "Devis NTR 2026(calculs auto.)"

// codeRegexp extracts the prestation code (e.g. "704NTR") at the end of a
// prestation row's label, such as:
//
//	"Frais d'admission ... code 704NTR"
var codeRegexp = regexp.MustCompile(`code\s+([0-9A-Za-z]+NTR)\s*$`)

// GenerateXLSX opens the template workbook, fills in the devis data,
// computes the prestation quantities/totals, then saves the result to
// outPath. The template file itself is never modified (SaveAs).
func GenerateXLSX(d Devis, templatePath, outPath string) error {
	f, err := excelize.OpenFile(templatePath)
	if err != nil {
		return fmt.Errorf("opening template: %w", err)
	}
	defer func() { _ = f.Close() }()

	if err := fillInformationsGenerales(f, d); err != nil {
		return err
	}

	if err := adaptImportantFields(f); err != nil {
		return err
	}
	if err := preparePrestationsLayout(f); err != nil {
		return err
	}

	total, err := fillPrestations(f, d)
	if err != nil {
		return err
	}

	// Order total (merged cell G25:H25).
	if err := f.SetCellValue(SheetName, "G25", total); err != nil {
		return err
	}
	// Keep the template's number format and borders, then make the amount
	// stand out. Replacing the style would remove the cell border.
	if err := setLargeBoldCell(f, "G25", 18); err != nil {
		return err
	}

	// Keep only the filled-in devis sheet: the "vierge" (blank) and
	// "FORFAITS" tabs from the template workbook must not show up in the
	// generated PDF (PDF export produces one page per sheet).
	if err := keepOnlySheet(f, SheetName); err != nil {
		return err
	}

	if err := f.SaveAs(outPath); err != nil {
		return fmt.Errorf("saving devis: %w", err)
	}
	return nil
}

// keepOnlySheet removes every sheet from the workbook except the one given,
// so that the PDF conversion produces a single page.
func keepOnlySheet(f *excelize.File, sheetToKeep string) error {
	for _, name := range f.GetSheetList() {
		if name == sheetToKeep {
			continue
		}
		if err := f.DeleteSheet(name); err != nil {
			return fmt.Errorf("deleting sheet %q: %w", name, err)
		}
	}
	f.SetActiveSheet(0)
	return nil
}

// adaptImportantFields keeps the name, first name and funeral operator
// inside their frames. Their font always remains the same size; only the
// field height grows when a long value needs additional lines.
func adaptImportantFields(f *excelize.File) error {
	// The operator area is visually a single field across E:H, but was not
	// merged in the source template. Merging it prevents text flowing over the
	// neighbouring cells when a company name is long.
	if err := f.MergeCell(SheetName, "E5", "H5"); err != nil {
		return fmt.Errorf("merging funeral operator field: %w", err)
	}

	rowHeights := map[int]float64{}
	for _, field := range []struct {
		cell     string
		row      int
		lineSize int
	}{
		{cell: "A5", row: 5, lineSize: 34},
		{cell: "A9", row: 9, lineSize: 34},
		{cell: "E5", row: 5, lineSize: 34},
	} {
		if err := setAdaptiveHeaderCell(f, field.cell); err != nil {
			return err
		}
		value, err := f.GetCellValue(SheetName, field.cell)
		if err != nil {
			return err
		}
		height := headerRowHeight(value, field.lineSize)
		if height > rowHeights[field.row] {
			rowHeights[field.row] = height
		}
	}
	for row, height := range rowHeights {
		if err := f.SetRowHeight(SheetName, row, height); err != nil {
			return err
		}
	}
	return nil
}

func headerRowHeight(value string, charactersPerLine int) float64 {
	lines := math.Ceil(float64(len([]rune(value))) / float64(charactersPerLine))
	if lines < 1 {
		lines = 1
	}
	// 24 points per line gives the fixed 16 pt font enough room without
	// clipping its descenders when the document is converted to PDF.
	return math.Max(30, lines*24)
}

func setAdaptiveHeaderCell(f *excelize.File, cell string) error {
	styleID, err := f.GetCellStyle(SheetName, cell)
	if err != nil {
		return err
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		return err
	}

	styleCopy := *style
	if style.Font == nil {
		styleCopy.Font = &excelize.Font{}
	} else {
		fontCopy := *style.Font
		styleCopy.Font = &fontCopy
	}
	styleCopy.Font.Size = 16
	styleCopy.Font.Bold = true
	if style.Alignment == nil {
		styleCopy.Alignment = &excelize.Alignment{}
	} else {
		alignmentCopy := *style.Alignment
		styleCopy.Alignment = &alignmentCopy
	}
	styleCopy.Alignment.WrapText = true
	styleCopy.Alignment.Vertical = "center"

	newStyleID, err := f.NewStyle(&styleCopy)
	if err != nil {
		return err
	}
	return f.SetCellStyle(SheetName, cell, cell, newStyleID)
}

// setLargeBoldCell preserves a cell's existing style and only enlarges and
// bolds its font.
func setLargeBoldCell(f *excelize.File, cell string, size float64) error {
	styleID, err := f.GetCellStyle(SheetName, cell)
	if err != nil {
		return err
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		return err
	}

	// GetStyle can return a style shared by several cells in the template.
	// Copy it before changing the font, otherwise unrelated cells (including
	// long prestation labels) inherit the enlarged font as well.
	styleCopy := *style
	if style.Font == nil {
		styleCopy.Font = &excelize.Font{}
	} else {
		fontCopy := *style.Font
		styleCopy.Font = &fontCopy
	}
	styleCopy.Font.Size = size
	styleCopy.Font.Bold = true

	newStyleID, err := f.NewStyle(&styleCopy)
	if err != nil {
		return err
	}
	return f.SetCellStyle(SheetName, cell, cell, newStyleID)
}

// preparePrestationsLayout gives the prestation description enough room to
// wrap inside its bordered cell and moves each NTR code into its own
// right-aligned cell.
func preparePrestationsLayout(f *excelize.File) error {
	lignes, err := listPrestations(f)
	if err != nil {
		return err
	}

	// Keep the table's total width unchanged: D becomes part of the label area
	// and E is a compact, dedicated code column.
	if err := f.SetColWidth(SheetName, "D", "D", 22); err != nil {
		return err
	}
	if err := f.SetColWidth(SheetName, "E", "E", 12); err != nil {
		return err
	}

	for _, ligne := range lignes {
		row := ligne.Row
		labelCell := fmt.Sprintf("A%d", row)
		codeCell := fmt.Sprintf("E%d", row)

		if err := f.UnmergeCell(SheetName, labelCell, codeCell); err != nil {
			return fmt.Errorf("unmerging prestation row %d: %w", row, err)
		}
		if err := f.MergeCell(SheetName, labelCell, fmt.Sprintf("D%d", row)); err != nil {
			return fmt.Errorf("merging prestation label row %d: %w", row, err)
		}
		if err := f.SetCellValue(SheetName, labelCell, prestationLabel(ligne.Libelle)); err != nil {
			return err
		}
		if err := f.SetCellValue(SheetName, codeCell, ligne.Code); err != nil {
			return err
		}
		if err := setPrestationAlignment(f, labelCell, "left", true); err != nil {
			return err
		}
		if err := setPrestationAlignment(f, codeCell, "right", false); err != nil {
			return err
		}
		// Two lines fit comfortably, including the longest current labels, with
		// room left for a future wording change.
		if err := f.SetRowHeight(SheetName, row, 48); err != nil {
			return err
		}
	}

	return nil
}

func prestationLabel(value string) string {
	return strings.TrimSpace(codeRegexp.ReplaceAllString(value, ""))
}

// setPrestationAlignment changes only the alignment of a cell while
// preserving its existing font, fill, number format and borders.
func setPrestationAlignment(f *excelize.File, cell, horizontal string, wrapText bool) error {
	styleID, err := f.GetCellStyle(SheetName, cell)
	if err != nil {
		return err
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		return err
	}

	styleCopy := *style
	if style.Alignment == nil {
		styleCopy.Alignment = &excelize.Alignment{}
	} else {
		alignmentCopy := *style.Alignment
		styleCopy.Alignment = &alignmentCopy
	}
	styleCopy.Alignment.Horizontal = horizontal
	styleCopy.Alignment.Vertical = "center"
	styleCopy.Alignment.WrapText = wrapText

	newStyleID, err := f.NewStyle(&styleCopy)
	if err != nil {
		return err
	}
	return f.SetCellStyle(SheetName, cell, cell, newStyleID)
}

// fillInformationsGenerales fills in the header cells (identity, dates,
// measurements, funeral operator...).
//
// IMPORTANT: in the template, these fields are merged cells that already
// contain a printed label (e.g. A5 = "Nom "). There is no separate blank
// cell to fill in: on the paper document, you write by hand right after the
// label. We reproduce that principle by replacing the cell's content with
// "Label : value". If your template evolves and a real input cell is added
// next to the label, just update the cell constants below.
func fillInformationsGenerales(f *excelize.File, d Devis) error {
	set := func(cell, value string) error {
		return f.SetCellValue(SheetName, cell, value)
	}

	if err := set("A5", labelValue("Nom", d.Nom)); err != nil {
		return err
	}
	if err := set("E5", labelValue("Opérateur Funéraire / Ville", "\n"+d.Operateur)); err != nil {
		return err
	}
	if d.NomJeuneFille != "" {
		if err := set("A7", labelValue("Nom de jeune fille", d.NomJeuneFille)); err != nil {
			return err
		}
	}
	if err := set("G8", labelValue("Taille du défunt", withUnit(d.TailleDefunt, "CM"))); err != nil {
		return err
	}
	if err := set("A9", labelValue("Prénom", d.Prenom)); err != nil {
		return err
	}
	if err := set("E9", "NANTERRE Le "+d.DateCommande); err != nil {
		return err
	}
	if err := set("G9", labelValue("Epaulement", withUnit(d.Epaulement, "CM"))); err != nil {
		return err
	}
	if err := set("G10", labelValue("Coude à coude", withUnit(d.CoudeACoude, "CM"))); err != nil {
		return err
	}
	if err := set("A11", fmt.Sprintf("Né(e) le %s   Décédé(e) le %s à %s",
		orTiret(d.DateNaissance), orTiret(d.DateDeces), orTiret(d.HeureDeces))); err != nil {
		return err
	}
	if err := set("G11", labelValue("Epaisseur", withUnit(d.Epaisseur, "CM"))); err != nil {
		return err
	}
	if err := set("A12", fmt.Sprintf("Date & Heure d'admission %s à %s",
		orTiret(d.DateAdmission), orTiret(d.HeureAdmission))); err != nil {
		return err
	}
	if err := set("E12", labelValue("Ville de décès", d.VilleDeces)); err != nil {
		return err
	}
	if err := set("A13", fmt.Sprintf("Date & Heure de départ %s à %s",
		orTiret(d.DateDepart), orTiret(d.HeureDepart))); err != nil {
		return err
	}
	if err := set("E13", labelValue("Code postal", d.CodePostal)); err != nil {
		return err
	}

	return nil
}

// fillPrestations walks the prestations table (rows 16 to 24), identifies
// each row by its code (e.g. "704NTR") rather than by its full label (more
// robust: the label can be long/reworded, while the code stays stable),
// fills in the quantity in column G and computes the total in column H
// (unit price incl. tax x quantity). Returns the overall order total.
func fillPrestations(f *excelize.File, d Devis) (float64, error) {
	lignes, err := listPrestations(f)
	if err != nil {
		return 0, err
	}

	found := map[string]bool{}
	var total float64

	for _, ligne := range lignes {
		qte, ok := d.Prestations[ligne.Code]
		if !ok || qte == 0 {
			continue
		}
		found[ligne.Code] = true

		gCell := fmt.Sprintf("G%d", ligne.Row)
		hCell := fmt.Sprintf("H%d", ligne.Row)

		if err := f.SetCellValue(SheetName, gCell, qte); err != nil {
			return 0, err
		}
		if err := setLargeBoldCell(f, gCell, 16); err != nil {
			return 0, err
		}
		amount := ligne.PrixTTC * qte
		if err := f.SetCellValue(SheetName, hCell, amount); err != nil {
			return 0, err
		}
		total += amount
	}

	// Flag any prestation codes sent by the frontend that don't exist in the template.
	for code := range d.Prestations {
		if !found[code] && d.Prestations[code] != 0 {
			return 0, fmt.Errorf("unknown prestation code in template: %s", code)
		}
	}

	return total, nil
}

// listPrestations scans column A of the table rows (from the row after the
// "PRESTATIONS" header down to the "Total de la commande" row) and extracts
// the code, label and unit price incl. tax (column F) for each row.
func listPrestations(f *excelize.File) ([]PrestationLigne, error) {
	rows, err := f.GetRows(SheetName)
	if err != nil {
		return nil, err
	}

	var lignes []PrestationLigne
	for i, row := range rows {
		rowNum := i + 1
		if len(row) == 0 {
			continue
		}
		libelle := strings.TrimSpace(row[0])
		matches := codeRegexp.FindStringSubmatch(libelle)
		code := ""
		if matches != nil {
			code = matches[1]
		} else {
			// Generated workbooks place the code in column E so the
			// description can use the full A:D width.
			code, err = f.GetCellValue(SheetName, fmt.Sprintf("E%d", rowNum))
			if err != nil {
				return nil, fmt.Errorf("reading prestation code E%d: %w", rowNum, err)
			}
			code = strings.TrimSpace(code)
			if !strings.HasSuffix(code, "NTR") {
				continue
			}
		}

		// Re-read cell F as a raw value (RawCellValue) instead of from
		// `row`, which contains the text already formatted for display
		// (e.g. "137.38 €"), in order to get a usable number.
		raw, err := f.GetCellValue(SheetName, fmt.Sprintf("F%d", rowNum), excelize.Options{RawCellValue: true})
		if err != nil {
			return nil, fmt.Errorf("reading unit price F%d: %w", rowNum, err)
		}
		prix, _ := strconv.ParseFloat(strings.ReplaceAll(strings.TrimSpace(raw), ",", "."), 64)

		lignes = append(lignes, PrestationLigne{
			Code:    code,
			Libelle: libelle,
			Row:     rowNum,
			PrixTTC: prix,
		})
	}
	return lignes, nil
}

// ListerPrestationsDisponibles is exposed for the API: lets the frontend
// dynamically fetch the list of prestations (code, label, price) as defined
// in the template, without duplicating that information client-side.
func ListerPrestationsDisponibles(templatePath string) ([]PrestationLigne, error) {
	f, err := excelize.OpenFile(templatePath)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()
	return listPrestations(f)
}

// --- Small formatting helpers ---

func labelValue(label, value string) string {
	if value == "" {
		return label
	}
	return label + " : " + value
}

func withUnit(value, unit string) string {
	if value == "" {
		return ""
	}
	return value + " " + unit
}

func orTiret(value string) string {
	if value == "" {
		return "….."
	}
	return value
}

