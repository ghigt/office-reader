package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

type Relationships struct {
	Items []Relationship `xml:"Relationship"`
}

type Relationship struct {
	Id     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// readFileFromZip reads a file from a ZIP archive by name.
// Returns nil, nil if the file is not found.
func readFileFromZip(r *zip.ReadCloser, name string) ([]byte, error) {
	for _, f := range r.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, nil
}

// readRelsFromZip parses a .rels file and returns a map of Id→Target
// for relationships whose Type ends with the given suffix (e.g. "hyperlink").
func readRelsFromZip(r *zip.ReadCloser, relsPath, typeSuffix string) (map[string]string, error) {
	data, err := readFileFromZip(r, relsPath)
	if err != nil {
		return nil, err
	}
	if data == nil {
		return make(map[string]string), nil
	}

	var relationships Relationships
	if err := xml.Unmarshal(data, &relationships); err != nil {
		return nil, fmt.Errorf("rels parsing error: %w", err)
	}

	rels := make(map[string]string)
	for _, rel := range relationships.Items {
		if strings.HasSuffix(rel.Type, "/"+typeSuffix) {
			rels[rel.Id] = rel.Target
		}
	}
	return rels, nil
}
