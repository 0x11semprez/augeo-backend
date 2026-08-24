package devis

// Devis represents the data submitted from the frontend to generate a
// "Maison Funéraire du Mont Valérien - Nanterre" order form / quote.
//
// All dates and times are strings already formatted on the frontend side
// (e.g. "01/07/2026" and "14h30") to keep things simple: the frontend
// controls the display format, the backend just inserts them into the
// document.
type Devis struct {
	// --- Deceased / family information ---
	Civilite     string `json:"civilite"`     // "M" or "Mme"
	Nom          string `json:"nom"`          // Last name
	NomNaissance string `json:"nomNaissance"` // Optional (birth name)
	Prenom       string `json:"prenom"`

	DateNaissance string `json:"dateNaissance"` // "DD/MM/YYYY"
	DateDeces     string `json:"dateDeces"`     // "DD/MM/YYYY"
	HeureDeces    string `json:"heureDeces"`    // "HHhMM"
	VilleDeces    string `json:"villeDeces"`
	CodePostal    string `json:"codePostal"`

	TailleDefunt string `json:"tailleDefunt"` // in cm, kept as text to accept "" when unknown
	Epaulement   string `json:"epaulement"`
	CoudeACoude  string `json:"coudeACoude"`
	Epaisseur    string `json:"epaisseur"`

	// --- Stay ---
	DateAdmission  string `json:"dateAdmission"`
	HeureAdmission string `json:"heureAdmission"`
	DateDepart     string `json:"dateDepart"`
	HeureDepart    string `json:"heureDepart"`

	// --- Funeral operator ---
	Operateur    string `json:"operateur"`    // Company name / city
	DateCommande string `json:"dateCommande"` // Order date ("NANTERRE Le ...")

	// --- Mentions (checkboxes shown at the top of the order form) ---
	PaiementChequeDepart bool `json:"paiementChequeDepart"`
	TresGrand            bool `json:"tresGrand"`
	ArriveeNuit          bool `json:"arriveeNuit"`

	// --- Metadata ---
	Initiales string `json:"initiales"` // Initials of the person who drafted the devis, printed at the bottom of the form

	// --- Prestations ---
	// Key = prestation code as printed on the quote (e.g. "704NTR", "756NTR", ...)
	// Value = ordered quantity. A missing or zero-quantity prestation is not billed.
	Prestations map[string]float64 `json:"prestations"`
}

// PrestationLigne describes one row of the prestations table as found
// dynamically in the template workbook (see generator.go).
type PrestationLigne struct {
	Code    string  // e.g. "704NTR"
	Libelle string  // full text of column A
	Row     int     // Excel row number (1-indexed)
	PrixTTC float64 // value found in column F
}
