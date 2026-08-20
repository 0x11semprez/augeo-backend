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

// printScale replaces the template's original 61% print scale. The template
// only fills a fraction of an A4 page at 100%, leaving large blank margins
// top and bottom once converted to PDF; printing larger uses that space
// instead of wasting it. Value picked empirically so the sheet still fits on
// a single page.
var printScale uint = 90

// codeRegexp extracts the prestation code (e.g. "704NTR") at the end of a
// prestation row's label, such as:
//
//	"Frais d'admission ... code 704NTR"
var codeRegexp = regexp.MustCompile(`code\s+([0-9A-Za-z]+NTR)\s*$`)

// prestationCodeRegexp matches a real prestation code (e.g. "704NTR"): one
// or more digits followed by "NTR". It excludes look-alike text such as the
// "code DSPMFNTR" column header, which also ends in "NTR" but isn't a code.
var prestationCodeRegexp = regexp.MustCompile(`^\d+NTR$`)

// GenerateXLSX opens the template workbook, fills in the devis data,
// computes the prestation quantities/totals, then saves the result to
// outPath. The template file itself is never modified (SaveAs).
func GenerateXLSX(d Devis, templatePath, outPath string) error {
	f, err := excelize.OpenFile(templatePath)
	if err != nil {
		return fmt.Errorf("opening template: %w", err)
	}
	defer func() { _ = f.Close() }()

	// Make room for the mentions banner above the "BON DE COMMANDE..."
	// letterhead box: insert a blank row at the very top and shift
	// everything else (including that box) down by one. Every cell
	// reference below this point uses the post-insertion row numbers.
	if err := f.InsertRows(SheetName, 1, 1); err != nil {
		return fmt.Errorf("inserting mentions row: %w", err)
	}
	// excelize treats a range starting exactly at the inserted row as
	// content that moved down, so the template's print area ($A$1:$I$35)
	// becomes $A$2:$I$36 instead of growing to include the new row. Put it
	// back to $A$1 so the mentions banner isn't clipped from the PDF export.
	if err := f.DeleteDefinedName(&excelize.DefinedName{Name: "_xlnm.Print_Area", Scope: SheetName}); err != nil {
		return fmt.Errorf("clearing print area: %w", err)
	}
	if err := f.SetDefinedName(&excelize.DefinedName{
		Name:     "_xlnm.Print_Area",
		RefersTo: fmt.Sprintf("'%s'!$A$1:$I$36", SheetName),
		Scope:    SheetName,
	}); err != nil {
		return fmt.Errorf("setting print area: %w", err)
	}

	if err := fillInformationsGenerales(f, d); err != nil {
		return err
	}
	if err := fillMentions(f, d); err != nil {
		return err
	}

	if err := adaptImportantFields(f, d); err != nil {
		return err
	}
	if err := adaptSecondaryFields(f); err != nil {
		return err
	}
	if err := repositionCodeHeader(f); err != nil {
		return err
	}
	if err := preparePrestationsLayout(f); err != nil {
		return err
	}

	total, err := fillPrestations(f, d)
	if err != nil {
		return err
	}

	// Order total (merged cell G26:H26).
	if err := f.SetCellValue(SheetName, "G26", total); err != nil {
		return err
	}
	// Keep the template's number format and borders, then make the amount
	// stand out. Replacing the style would remove the cell border.
	if err := setLargeBoldCell(f, "G26", 18); err != nil {
		return err
	}

	// Keep only the filled-in devis sheet: the "vierge" (blank) and
	// "FORFAITS" tabs from the template workbook must not show up in the
	// generated PDF (PDF export produces one page per sheet).
	if err := keepOnlySheet(f, SheetName); err != nil {
		return err
	}

	if err := f.SetPageLayout(SheetName, &excelize.PageLayoutOptions{AdjustTo: &printScale}); err != nil {
		return fmt.Errorf("setting print scale: %w", err)
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

// adaptImportantFields keeps the name, first name and funeral operator name
// inside their frames, and makes the actual values (the variables a reader
// scans for first) noticeably larger than their printed labels - the labels
// stay at the same size as every other label on the form. Only the field
// height grows when a long value needs additional lines.
func adaptImportantFields(f *excelize.File, d Devis) error {
	// The operator area is visually a single bordered box spanning E6:H8 in
	// the template, but only E6 itself was a real cell. It holds two
	// distinct pieces of text that must not share one alignment: the label
	// ("Opérateur Funéraire / Ville :", left-aligned like the other labels,
	// filled in by fillInformationsGenerales) on row 6, and the operator's
	// name (centered, the only part that should be) on row 7. Merging both
	// rows across E:H prevents text flowing over the neighbouring cells when
	// the label or the company name is long.
	if err := f.MergeCell(SheetName, "E6", "H6"); err != nil {
		return fmt.Errorf("merging funeral operator label: %w", err)
	}
	if err := f.MergeCell(SheetName, "E7", "H7"); err != nil {
		return fmt.Errorf("merging funeral operator name: %w", err)
	}

	// labelFontSize matches the template's original size for every field
	// label on the form: only the value that follows a label stands out.
	// identityFontSize applies to Nom, Prénom and Nom de naissance alike so
	// the three sit at the same size instead of Nom de naissance looking
	// undersized next to the other two.
	const labelFontSize = 11
	const identityFontSize = 20
	const operatorFontSize = 16

	rowHeights := map[int]float64{}
	for _, field := range []struct {
		cell             string
		row              int
		startCol, endCol string
		horizontal       string
		fontSize         float64
		bold             bool   // only matters when label == "": setSplitSizeCell controls its own runs' weight
		label, value     string // when set, cell renders as "label : value" with label at labelFontSize and value at fontSize
	}{
		{cell: "A6", row: 6, startCol: "A", endCol: "C", fontSize: identityFontSize, label: "Nom", value: d.Nom},
		{cell: "A8", row: 8, startCol: "A", endCol: "C", fontSize: identityFontSize, label: "Nom de naissance", value: d.NomNaissance},
		{cell: "A10", row: 10, startCol: "A", endCol: "C", fontSize: identityFontSize, label: "Prénom", value: d.Prenom},
		{cell: "E7", row: 7, startCol: "E", endCol: "H", horizontal: "center", fontSize: operatorFontSize, bold: false},
	} {
		if err := setAdaptiveHeaderCell(f, field.cell, field.fontSize, field.horizontal, field.bold); err != nil {
			return err
		}
		if field.label != "" {
			if err := setSplitSizeCell(f, field.cell, field.label, field.value, labelFontSize, field.fontSize); err != nil {
				return err
			}
		}
		value, err := f.GetCellValue(SheetName, field.cell)
		if err != nil {
			return err
		}
		charsPerLine, err := approxCharsPerLine(f, field.startCol, field.endCol, field.fontSize)
		if err != nil {
			return err
		}
		height := headerRowHeight(value, charsPerLine)
		if height > rowHeights[field.row] {
			rowHeights[field.row] = height
		}
	}
	for row, height := range rowHeights {
		if err := f.SetRowHeight(SheetName, row, height); err != nil {
			return err
		}
	}

	// The funeral operator's label (E6, filled in by fillInformationsGenerales)
	// gets the same bold treatment as Nom/Prénom/Nom de naissance's labels,
	// while its value (E7, just handled above) stays normal weight.
	if err := setBoldCell(f, "E6"); err != nil {
		return fmt.Errorf("bolding funeral operator label: %w", err)
	}
	return nil
}

// setSplitSizeCell renders a cell as "label : value" using two rich-text
// runs so the label and the value can each have their own font size and
// weight - a single cell's alignment applies uniformly, but rich-text runs
// each carry their own font. The label is bold (it's the fixed part a reader
// recognizes at a glance); the value stays normal weight.
func setSplitSizeCell(f *excelize.File, cell, label, value string, labelSize, valueSize float64) error {
	return f.SetCellRichText(SheetName, cell, []excelize.RichTextRun{
		{Text: label + " : ", Font: &excelize.Font{Family: "Calibri", Size: labelSize, Bold: true}},
		{Text: value, Font: &excelize.Font{Family: "Calibri", Size: valueSize}},
	})
}

// adaptSecondaryFields prevents the funeral operator's label, the stay
// dates/times and the body measurements from being cropped when a value is
// unusually long. Unlike adaptImportantFields, it doesn't change font size,
// weight or alignment - it only turns on wrapping and grows the row height
// to fit, so short values (the common case) render exactly as before.
func adaptSecondaryFields(f *excelize.File) error {
	// A12/A13/A14 aren't merged across their usable width in the template
	// (unlike Nom/Prénom): merge them so wrapping uses the real available
	// width (up to column D, where the measurements column starts) instead
	// of column A alone.
	for _, row := range []int{12, 13, 14} {
		cell := fmt.Sprintf("A%d", row)
		endCell := fmt.Sprintf("D%d", row)
		if err := f.MergeCell(SheetName, cell, endCell); err != nil {
			return fmt.Errorf("merging row %d date field: %w", row, err)
		}
	}

	fields := []struct {
		cell             string
		row              int
		startCol, endCol string
	}{
		{cell: "E6", row: 6, startCol: "E", endCol: "H"},
		{cell: "A12", row: 12, startCol: "A", endCol: "D"},
		{cell: "A13", row: 13, startCol: "A", endCol: "D"},
		{cell: "A14", row: 14, startCol: "A", endCol: "D"},
		{cell: "G9", row: 9, startCol: "G", endCol: "H"},
		{cell: "G10", row: 10, startCol: "G", endCol: "H"},
		{cell: "G11", row: 11, startCol: "G", endCol: "H"},
		{cell: "G12", row: 12, startCol: "G", endCol: "H"},
	}

	for _, field := range fields {
		fontSize, err := enableWrap(f, field.cell)
		if err != nil {
			return err
		}

		value, err := f.GetCellValue(SheetName, field.cell)
		if err != nil {
			return err
		}
		charsPerLine, err := approxCharsPerLine(f, field.startCol, field.endCol, fontSize)
		if err != nil {
			return err
		}
		needed := headerRowHeight(value, charsPerLine)

		current, err := f.GetRowHeight(SheetName, field.row)
		if err != nil {
			return err
		}
		if needed > current {
			if err := f.SetRowHeight(SheetName, field.row, needed); err != nil {
				return err
			}
		}
	}
	return nil
}

// enableWrap turns on word-wrap and vertical centering for cell while
// preserving its existing font, size, weight and horizontal alignment. It
// returns the cell's font size, needed by the caller to size the row.
func enableWrap(f *excelize.File, cell string) (float64, error) {
	styleID, err := f.GetCellStyle(SheetName, cell)
	if err != nil {
		return 0, err
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		return 0, err
	}

	fontSize := 11.0
	if style.Font != nil && style.Font.Size > 0 {
		fontSize = style.Font.Size
	}

	styleCopy := *style
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
		return 0, err
	}
	return fontSize, f.SetCellStyle(SheetName, cell, cell, newStyleID)
}

// approxCharsPerLine estimates how many characters at fontSize fit on one
// line of a merged cell range, from the workbook's actual column widths for
// that range. This ties the line-wrap decision to the real case size instead
// of a fixed character count, so it keeps working if the template's columns
// are ever resized.
func approxCharsPerLine(f *excelize.File, startCol, endCol string, fontSize float64) (float64, error) {
	startIdx, err := excelize.ColumnNameToNumber(startCol)
	if err != nil {
		return 0, err
	}
	endIdx, err := excelize.ColumnNameToNumber(endCol)
	if err != nil {
		return 0, err
	}

	var totalWidth float64
	for i := startIdx; i <= endIdx; i++ {
		name, err := excelize.ColumnNumberToName(i)
		if err != nil {
			return 0, err
		}
		width, err := f.GetColWidth(SheetName, name)
		if err != nil {
			return 0, err
		}
		totalWidth += width
	}

	// An Excel column-width unit is calibrated to the width of the "0"
	// character in the workbook's default font (11pt). Bold characters at a
	// larger font take up proportionally more horizontal space, so scale the
	// usable character count down accordingly.
	const referenceFontSize = 11.0
	const boldWidthFactor = 0.85
	chars := totalWidth * (referenceFontSize / fontSize) * boldWidthFactor
	if chars < 1 {
		chars = 1
	}
	return chars, nil
}

func headerRowHeight(value string, charsPerLine float64) float64 {
	lines := math.Ceil(float64(len([]rune(value))) / charsPerLine)
	if lines < 1 {
		lines = 1
	}
	if lines > 1 {
		// approxCharsPerLine is a character-count estimate, but real
		// word-wrap breaks at spaces/hyphens, not at an exact character
		// count: a long hyphenated word landing just past the boundary
		// forces an early break, so actual wrapped text can need one more
		// line than this estimate predicts. Add that margin whenever
		// wrapping already kicks in - single-line values (the vast
		// majority) are unaffected, so short values render exactly as
		// before.
		lines++
	}
	// 24 points per line gives the fixed 16 pt font enough room without
	// clipping its descenders when the document is converted to PDF.
	return math.Max(30, lines*24)
}

func setAdaptiveHeaderCell(f *excelize.File, cell string, fontSize float64, horizontal string, bold bool) error {
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
	styleCopy.Font.Size = fontSize
	styleCopy.Font.Bold = bold
	if style.Alignment == nil {
		styleCopy.Alignment = &excelize.Alignment{}
	} else {
		alignmentCopy := *style.Alignment
		styleCopy.Alignment = &alignmentCopy
	}
	styleCopy.Alignment.WrapText = true
	styleCopy.Alignment.Vertical = "center"
	if horizontal != "" {
		styleCopy.Alignment.Horizontal = horizontal
	}

	newStyleID, err := f.NewStyle(&styleCopy)
	if err != nil {
		return err
	}
	return f.SetCellStyle(SheetName, cell, cell, newStyleID)
}

// setBoldCell preserves a cell's existing style (size, alignment, borders,
// wrapping) and only makes its font bold.
func setBoldCell(f *excelize.File, cell string) error {
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
	styleCopy.Font.Bold = true

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
		if err := setPrestationAlignment(f, labelCell, "left", "center", true); err != nil {
			return err
		}
		if err := setPrestationAlignment(f, codeCell, "right", "center", false); err != nil {
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

// repositionCodeHeader moves the "DSPMFNTR" label out of the narrow rotated
// column (I16:I25, running the full height of the prestations table) and
// turns it into a normal horizontal header sitting right above the codes
// column (E), alongside "Prix unitaire TTC", "Qté" and "Total".
func repositionCodeHeader(f *excelize.File) error {
	if err := f.UnmergeCell(SheetName, "A16", "E16"); err != nil {
		return fmt.Errorf("unmerging prestations header: %w", err)
	}
	if err := f.MergeCell(SheetName, "A16", "D16"); err != nil {
		return fmt.Errorf("re-merging prestations header: %w", err)
	}
	if err := f.SetCellValue(SheetName, "E16", "DSPMFNTR"); err != nil {
		return err
	}
	if err := setLargeBoldCell(f, "E16", 11); err != nil {
		return err
	}
	if err := setPrestationAlignment(f, "E16", "center", "center", true); err != nil {
		return err
	}

	if err := f.UnmergeCell(SheetName, "I16", "I25"); err != nil {
		return fmt.Errorf("unmerging leftover code label column: %w", err)
	}
	return f.SetCellValue(SheetName, "I16", "")
}

// mentionsEmptyRowHeight keeps the banner row unobtrusive when no mention is
// checked, instead of leaving a visible blank bar above the letterhead.
const mentionsEmptyRowHeight = 15.75

// mentionsFontSize and mentionsRowHeight make the checked mentions stand out
// (large, bold) while leaving visible breathing room between them and the
// "BON DE COMMANDE..." letterhead box right below.
const (
	mentionsFontSize  = 18
	mentionsRowHeight = 36
)

// fillMentions writes the checkbox mentions banner in row 1, the blank row
// GenerateXLSX inserts above the "BON DE COMMANDE..." letterhead box, so
// it's the very first thing a reader sees - above the letterhead, not below
// it. Only mentions the frontend actually checked are printed: an unchecked
// mention is omitted entirely rather than shown with an empty box.
func fillMentions(f *excelize.File, d Devis) error {
	mentions := []struct {
		label   string
		checked bool
	}{
		{"Paiement par chèque au départ", d.PaiementChequeDepart},
		{"Très grand", d.TresGrand},
		{"Arrivée de nuit", d.ArriveeNuit},
	}

	var parts []string
	for _, m := range mentions {
		if !m.checked {
			continue
		}
		parts = append(parts, "☒ "+m.label)
	}

	if err := f.MergeCell(SheetName, "A1", "H1"); err != nil {
		return fmt.Errorf("merging mentions banner: %w", err)
	}
	if len(parts) == 0 {
		if err := f.SetCellValue(SheetName, "A1", ""); err != nil {
			return err
		}
		return f.SetRowHeight(SheetName, 1, mentionsEmptyRowHeight)
	}

	if err := f.SetCellValue(SheetName, "A1", strings.Join(parts, "     ")); err != nil {
		return err
	}
	if err := setLargeBoldCell(f, "A1", mentionsFontSize); err != nil {
		return err
	}
	// Top-aligned in a taller row so the text sits close to the top of the
	// page, leaving visible space below it before the letterhead box starts.
	if err := setPrestationAlignment(f, "A1", "center", "top", false); err != nil {
		return err
	}
	return f.SetRowHeight(SheetName, 1, mentionsRowHeight)
}

// setPrestationAlignment changes only the alignment of a cell while
// preserving its existing font, fill, number format and borders.
func setPrestationAlignment(f *excelize.File, cell, horizontal, vertical string, wrapText bool) error {
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
	styleCopy.Alignment.Vertical = vertical
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

	if err := set("A6", labelValue("Nom", d.Nom)); err != nil {
		return err
	}
	if err := set("E6", "Opérateur Funéraire / Ville :"); err != nil {
		return err
	}
	if err := set("E7", d.Operateur); err != nil {
		return err
	}
	// A8 ("Nom de naissance") is left unset here: adaptImportantFields fills
	// it with the same split-size treatment as Nom/Prénom so all three sit
	// at the same size.
	if err := set("G9", labelValue("Taille du défunt", withUnit(d.TailleDefunt, "CM"))); err != nil {
		return err
	}
	if err := set("A10", labelValue("Prénom", d.Prenom)); err != nil {
		return err
	}
	if err := set("E10", "NANTERRE Le "+d.DateCommande); err != nil {
		return err
	}
	if err := set("G10", labelValue("Epaulement", withUnit(d.Epaulement, "CM"))); err != nil {
		return err
	}
	if err := set("G11", labelValue("Coude à coude", withUnit(d.CoudeACoude, "CM"))); err != nil {
		return err
	}
	if err := set("A12", fmt.Sprintf("Né(e) le %s   Décédé(e) le %s à %s",
		orTiret(d.DateNaissance), orTiret(d.DateDeces), orTiret(d.HeureDeces))); err != nil {
		return err
	}
	if err := set("G12", labelValue("Epaisseur", withUnit(d.Epaisseur, "CM"))); err != nil {
		return err
	}
	if err := set("A13", fmt.Sprintf("Date & Heure d'admission %s à %s",
		orTiret(d.DateAdmission), orTiret(d.HeureAdmission))); err != nil {
		return err
	}
	if err := set("E13", labelValue("Ville de décès", d.VilleDeces)); err != nil {
		return err
	}
	if err := set("A14", fmt.Sprintf("Date & Heure de départ %s à %s",
		orTiret(d.DateDepart), orTiret(d.HeureDepart))); err != nil {
		return err
	}
	if err := set("E14", labelValue("Code postal", d.CodePostal)); err != nil {
		return err
	}

	return nil
}

// fillPrestations walks the prestations table (rows 17 to 25), identifies
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
			if !prestationCodeRegexp.MatchString(code) {
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
