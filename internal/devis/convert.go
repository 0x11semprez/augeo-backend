package devis

import (
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"sync"
	"time"
)

// soffice doesn't handle concurrent conversions well when they share the
// same user profile (lock files). We therefore serialize conversions with a
// global mutex. For high-volume usage, replace this approach with a
// dedicated conversion service pool (e.g. Gotenberg) that natively handles
// concurrency.
var conversionMutex sync.Mutex

// ConvertToPDF converts an xlsx file to PDF using headless LibreOffice and
// returns the path to the generated PDF.
//
// Requirement: LibreOffice must be installed on the machine running the
// server (Debian/Ubuntu package: `libreoffice-calc`).
func ConvertToPDF(xlsxPath, outDir string) (string, error) {
	conversionMutex.Lock()
	defer conversionMutex.Unlock()

	if err := os.MkdirAll(outDir, 0o755); err != nil {
		return "", fmt.Errorf("creating output directory: %w", err)
	}

	// Isolated, temporary LibreOffice user profile to avoid conflicts with
	// any other instances running on the machine.
	profileDir, err := os.MkdirTemp("", "lo-profile-*")
	if err != nil {
		return "", fmt.Errorf("creating temporary profile: %w", err)
	}
	defer func() { _ = os.RemoveAll(profileDir) }()

	ctx, cancel := context.WithTimeout(context.Background(), 60*time.Second)
	defer cancel()

	cmd := exec.CommandContext(
		ctx, "soffice",
		"--headless",
		"--norestore",
		"--convert-to", "pdf",
		"--outdir", outDir,
		"-env:UserInstallation=file://"+profileDir,
		xlsxPath,
	)

	output, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("LibreOffice conversion failed: %w (output: %s)", err, string(output))
	}

	base := filepath.Base(xlsxPath)
	pdfName := base[:len(base)-len(filepath.Ext(base))] + ".pdf"
	pdfPath := filepath.Join(outDir, pdfName)

	if _, err := os.Stat(pdfPath); err != nil {
		return "", fmt.Errorf("expected PDF not found after conversion (%s): %w", pdfPath, err)
	}

	return pdfPath, nil
}
