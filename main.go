package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Fprintln(os.Stderr, "Usage: office-reader <file.docx|.pptx|.xlsx>")
		os.Exit(1)
	}
	path := os.Args[1]
	ext := strings.ToLower(filepath.Ext(path))

	var lines []string
	var err error

	switch ext {
	case ".docx":
		lines, err = readDocx(path)
	case ".pptx":
		lines, err = readPptx(path)
	case ".xlsx":
		lines, err = readXlsx(path)
	default:
		fmt.Fprintf(os.Stderr, "Unsupported format: %s\n", ext)
		os.Exit(1)
	}

	if err != nil {
		fmt.Fprintln(os.Stderr, "Error:", err)
		os.Exit(1)
	}

	for i, line := range lines {
		fmt.Printf("[%d] %s\n", i, line)
	}
}
