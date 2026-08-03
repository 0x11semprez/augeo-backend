package devis

import (
	"fmt"
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
	defer f.Close()

	if err := fillInformationsGenerales(f, d); err != nil {
		return err
	}

	if err := emphasizeNomEtOperateur(f); err != nil {
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
	_ = f.SetCellStyle(SheetName, "G25", "G25", mustStyleID(f, "#,##0.00 \"€\""))

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

// emphasizeNomEtOperateur enlarges the font of the "Nom" and "Opérateur
// Funéraire / Ville" fields to make them stand out on the document, while
// keeping the original font and bold style.
func emphasizeNomEtOperateur(f *excelize.File) error {
	const largeSize = 16

	for _, cell := range []string{"A5", "E5"} {
		styleID, err := f.GetCellStyle(SheetName, cell)
		if err != nil {
			return err
		}
		style, err := f.GetStyle(styleID)
		if err != nil {
			return err
		}
		if style.Font == nil {
			style.Font = &excelize.Font{}
		}
		style.Font.Size = largeSize
		style.Font.Bold = true

		newStyleID, err := f.NewStyle(style)
		if err != nil {
			return err
		}
		if err := f.SetCellStyle(SheetName, cell, cell, newStyleID); err != nil {
			return err
		}
	}

	// Row 5 now holds bigger text, so we increase its height slightly to
	// give the larger font some room (avoids visual clipping).
	if err := f.SetRowHeight(SheetName, 5, 30); err != nil {
		return err
	}

	return nil
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
	if err := set("E5", labelValue("Opérateur Funéraire / Ville", d.Operateur)); err != nil {
		return err
	}
	if err := set("A6", civilite(d.Civilite)); err != nil {
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
		if matches == nil {
			continue
		}
		code := matches[1]

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
	defer f.Close()
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

func civilite(c string) string {
	c = strings.TrimSpace(strings.ToUpper(c))
	switch c {
	case "M", "MONSIEUR":
		return "M."
	case "MME", "MADAME":
		return "Mme"
	default:
		return "M / Mme"
	}
}

// mustStyleID creates (or retrieves) a numeric style and returns its ID.
// On error, it returns the default style (0) rather than failing the whole
// generation over a simple display-format issue.
func mustStyleID(f *excelize.File, numFmt string) int {
	style, err := f.NewStyle(&excelize.Style{CustomNumFmt: &numFmt})
	if err != nil {
		return 0
	}
	return style
}
