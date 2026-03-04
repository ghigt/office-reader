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

type ContentTypes struct {
	Overrides []ContentTypeOverride `xml:"Override"`
}

type ContentTypeOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// findMainDocumentPath reads [Content_Types].xml and returns the path of the
// main wordprocessingml document part (e.g. "word/document.xml" or "word/document2.xml").
func findMainDocumentPath(r *zip.ReadCloser) (string, error) {
	data, err := readFileFromZip(r, "[Content_Types].xml")
	if err != nil {
		return "", err
	}
	if data == nil {
		return "", fmt.Errorf("[Content_Types].xml not found")
	}

	var ct ContentTypes
	if err := xml.Unmarshal(data, &ct); err != nil {
		return "", fmt.Errorf("content types parsing error: %w", err)
	}

	for _, o := range ct.Overrides {
		if strings.Contains(o.ContentType, "wordprocessingml.document.main") {
			return strings.TrimPrefix(o.PartName, "/"), nil
		}
	}
	return "", fmt.Errorf("main document part not found in [Content_Types].xml")
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
