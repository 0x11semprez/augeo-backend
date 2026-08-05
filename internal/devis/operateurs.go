package devis

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

const (
	sheetOperateurs = "Liste 2024"
	sheetCommunes   = "CP VILLE"
)

// Operateur describes one funeral operator entry, used by the frontend to
// let the user search/select an operator and to pre-address the "open
// Outlook" email.
type Operateur struct {
	Nom       string `json:"nom"`
	Email     string `json:"email"`
	Telephone string `json:"telephone"`
}

var (
	emailRegexp = regexp.MustCompile(`^\S+@\S+\.\S+$`)
	phoneRegexp = regexp.MustCompile(`^0[0-9. ]{8,}$`)
)

// ListerOperateurs reads the "Liste 2024" sheet of the funeral operators
// workbook. The operator name is reliably in column B throughout the sheet,
// but later rows were appended with the email/phone shifted into different
// columns than the header suggests, so email and phone are found by
// pattern-matching each cell in the row instead of trusting a fixed index.
func ListerOperateurs(path string) ([]Operateur, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("opening operators workbook: %w", err)
	}
	defer f.Close()

	rows, err := f.GetRows(sheetOperateurs)
	if err != nil {
		return nil, fmt.Errorf("reading sheet %q: %w", sheetOperateurs, err)
	}

	var operateurs []Operateur
	for i, row := range rows {
		if i == 0 || len(row) < 2 {
			continue // header row / blank separator row
		}
		nom := strings.TrimSpace(row[1])
		if nom == "" {
			continue
		}

		var email, telephone string
		for j, cell := range row {
			if j == 1 {
				continue // already used as the name
			}
			value := strings.TrimSpace(cell)
			switch {
			case value == "":
				continue
			case email == "" && emailRegexp.MatchString(value):
				email = value
			case telephone == "" && phoneRegexp.MatchString(value):
				telephone = value
			}
		}

		operateurs = append(operateurs, Operateur{Nom: nom, Email: email, Telephone: telephone})
	}
	return operateurs, nil
}

// ChargerCommunesParCodePostal reads the "CP VILLE" sheet and returns a
// lookup from postal code to the commune name(s) sharing it (a French
// postal code often covers several communes).
func ChargerCommunesParCodePostal(path string) (map[string][]string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("opening operators workbook: %w", err)
	}
	defer f.Close()

	rows, err := f.GetRows(sheetCommunes)
	if err != nil {
		return nil, fmt.Errorf("reading sheet %q: %w", sheetCommunes, err)
	}

	communes := make(map[string][]string)
	for i, row := range rows {
		if i == 0 || len(row) < 2 {
			continue
		}
		codePostal := strings.TrimSpace(row[0])
		nom := strings.TrimSpace(row[1])
		if codePostal == "" || nom == "" {
			continue
		}
		communes[codePostal] = append(communes[codePostal], nom)
	}
	return communes, nil
}
