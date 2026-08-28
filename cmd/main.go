// Command devis-api: HTTP server that generates, from an Excel template, a
// "Maison Funéraire du Mont Valérien - Nanterre" quote/order form filled in
// with data sent by a frontend, and returns it as a PDF.
//
// Start with:
//
//	go run . -addr :8080 -template ./template/devis_template.xlsx
//
// System requirement: LibreOffice (libreoffice-calc package) must be
// installed for PDF conversion.
package main

import (
	"encoding/json"
	"flag"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"time"

	"devis-api/internal/devis"
	"devis-api/internal/generator"

	"github.com/google/uuid"
)

var (
	addr           = flag.String("addr", ":8080", "server listen address")
	templatePath   = flag.String("template", "./template/devis_template.xlsx", "path to the template workbook")
	operateursPath = flag.String("operateurs", "./data/operateurs.xlsx", "path to the funeral operators workbook")
	workDir        = flag.String("workdir", "./tmp", "working directory for generated files")
)

func main() {
	flag.Parse()

	if _, err := os.Stat(*templatePath); err != nil {
		log.Fatalf("template not found (%s): %v", *templatePath, err)
	}
	if err := os.MkdirAll(*workDir, 0o755); err != nil {
		log.Fatalf("could not create working directory: %v", err)
	}

	operateurs, err := devis.ListerOperateurs(*operateursPath)
	if err != nil {
		log.Fatalf("could not read funeral operators (%s): %v", *operateursPath, err)
	}
	communes, err := devis.ChargerCommunesParCodePostal(*operateursPath)
	if err != nil {
		log.Fatalf("could not read postal codes (%s): %v", *operateursPath, err)
	}

	mux := http.NewServeMux()
	mux.HandleFunc("POST /api/devis", withCORS(handleGenererDevis))
	mux.HandleFunc("OPTIONS /api/devis", withCORS(func(w http.ResponseWriter, r *http.Request) {}))
	mux.HandleFunc("GET /api/prestations", withCORS(handleListerPrestations))
	mux.HandleFunc("OPTIONS /api/prestations", withCORS(func(w http.ResponseWriter, r *http.Request) {}))
	mux.HandleFunc("GET /api/operateurs", withCORS(func(w http.ResponseWriter, r *http.Request) {
		writeJSON(w, http.StatusOK, operateurs)
	}))
	mux.HandleFunc("OPTIONS /api/operateurs", withCORS(func(w http.ResponseWriter, r *http.Request) {}))
	mux.HandleFunc("GET /api/villes", withCORS(func(w http.ResponseWriter, r *http.Request) {
		villes := communes[r.URL.Query().Get("codePostal")]
		if villes == nil {
			villes = []string{}
		}
		writeJSON(w, http.StatusOK, villes)
	}))
	mux.HandleFunc("OPTIONS /api/villes", withCORS(func(w http.ResponseWriter, r *http.Request) {}))
	mux.HandleFunc("GET /healthz", func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusOK)
		_, _ = w.Write([]byte("ok"))
	})

	log.Printf("devis-api server started on %s (template: %s, operateurs: %s)", *addr, *templatePath, *operateursPath)
	log.Fatal(http.ListenAndServe(*addr, mux))
}

// withCORS allows calls from a frontend served on a different origin
// (useful in development). Restrict this to your own domain in production.
func withCORS(next http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")
		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusNoContent)
			return
		}
		next(w, r)
	}
}

// handleListerPrestations returns the list of prestations available in the
// template (code, label, unit price incl. tax), so the frontend can build
// its form dynamically without duplicating this information.
func handleListerPrestations(w http.ResponseWriter, r *http.Request) {
	lignes, err := generator.ListerPrestationsDisponibles(*templatePath)
	if err != nil {
		writeError(w, http.StatusInternalServerError, "could not read prestations: "+err.Error())
		return
	}
	writeJSON(w, http.StatusOK, lignes)
}

// handleGenererDevis receives the devis data as JSON, generates the xlsx
// then converts it to PDF, and returns the PDF directly to the client.
// Add ?format=xlsx to the URL to get the raw Excel file instead of the PDF
// (useful for review/archiving).
func handleGenererDevis(w http.ResponseWriter, r *http.Request) {
	var d devis.Devis
	if err := json.NewDecoder(r.Body).Decode(&d); err != nil {
		writeError(w, http.StatusBadRequest, "invalid JSON: "+err.Error())
		return
	}
	if err := validate(d); err != nil {
		writeError(w, http.StatusBadRequest, err.Error())
		return
	}

	id := uuid.NewString()
	xlsxPath := filepath.Join(*workDir, id+".xlsx")

	if err := generator.GenerateXLSX(d, *templatePath, xlsxPath); err != nil {
		writeError(w, http.StatusInternalServerError, "could not generate devis: "+err.Error())
		return
	}
	defer func() { _ = os.Remove(xlsxPath) }()

	fileName := "devis_" + safe(d.Nom) + "_" + safe(d.Prenom)

	if r.URL.Query().Get("format") == "xlsx" {
		w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		w.Header().Set("Content-Disposition", `attachment; filename="`+fileName+`.xlsx"`)
		http.ServeFile(w, r, xlsxPath)
		return
	}

	pdfPath, err := devis.ConvertToPDF(xlsxPath, *workDir)
	if err != nil {
		writeError(w, http.StatusInternalServerError, "could not convert to PDF: "+err.Error())
		return
	}
	defer func() { _ = os.Remove(pdfPath) }()

	w.Header().Set("Content-Type", "application/pdf")
	w.Header().Set("Content-Disposition", `attachment; filename="`+fileName+`.pdf"`)
	http.ServeFile(w, r, pdfPath)
}

// validate performs minimal checks on required fields.
func validate(d devis.Devis) error {
	if d.Nom == "" || d.Prenom == "" {
		return errBadRequest("the deceased's last name and first name are required")
	}
	if len(d.Prestations) == 0 {
		return errBadRequest("at least one prestation must be provided")
	}
	return nil
}

type errBadRequest string

func (e errBadRequest) Error() string { return string(e) }

func writeJSON(w http.ResponseWriter, status int, v any) {
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(status)
	_ = json.NewEncoder(w).Encode(v)
}

func writeError(w http.ResponseWriter, status int, message string) {
	writeJSON(w, status, map[string]string{"error": message, "timestamp": time.Now().Format(time.RFC3339)})
}

func safe(s string) string {
	out := make([]rune, 0, len(s))
	for _, r := range s {
		switch {
		case r >= 'a' && r <= 'z', r >= 'A' && r <= 'Z', r >= '0' && r <= '9':
			out = append(out, r)
		default:
			out = append(out, '_')
		}
	}
	if len(out) == 0 {
		return "devis"
	}
	return string(out)
}
