package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"sort"
	"strconv"
	"strings"
)

// extractTextFromXML walks arbitrary PPTX slide XML and extracts all text.
// It handles shapes (sp), grouped shapes (grpSp), graphic frames (graphicFrame)
// which contain tables, and any nested combination of these.
func extractTextFromXML(d *xml.Decoder) []string {
	var lines []string
	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "sp", "grpSp", "spTree":
			// Recurse into shapes, group shapes, and shape trees
			lines = append(lines, extractTextFromXML(d)...)
		case "graphicFrame":
			// graphicFrame contains tables and other graphic objects
			lines = append(lines, extractGraphicFrame(d)...)
		case "txBody":
			lines = append(lines, extractTxBody(d)...)
		default:
			d.Skip()
		}
	}
	return lines
}

// extractGraphicFrame processes a graphicFrame element, looking for tables.
func extractGraphicFrame(d *xml.Decoder) []string {
	var lines []string
	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tbl":
				lines = append(lines, extractTable(d)...)
			default:
				// Recurse into graphic/graphicData wrappers
				lines = append(lines, extractGraphicFrame(d)...)
			}
		case xml.EndElement:
			return lines
		}
	}
	return lines
}

// extractTable processes a table element, extracting text from all cells.
func extractTable(d *xml.Decoder) []string {
	var lines []string
	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "tr" {
				row := extractTableRow(d)
				if row != "" {
					lines = append(lines, row)
				}
			} else {
				d.Skip()
			}
		case xml.EndElement:
			return lines
		}
	}
	return lines
}

// extractTableRow processes a table row, returning cells joined with " | ".
func extractTableRow(d *xml.Decoder) string {
	var cells []string
	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "tc" {
				cell := extractTableCell(d)
				cells = append(cells, cell)
			} else {
				d.Skip()
			}
		case xml.EndElement:
			if len(cells) > 0 {
				return "| " + strings.Join(cells, " | ") + " |"
			}
			return ""
		}
	}
	return ""
}

// extractTableCell processes a table cell, extracting text from txBody paragraphs.
func extractTableCell(d *xml.Decoder) string {
	var parts []string
	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "txBody" {
				parts = append(parts, extractTxBody(d)...)
			} else {
				d.Skip()
			}
		case xml.EndElement:
			return strings.Join(parts, " ")
		}
	}
	return strings.Join(parts, " ")
}

// extractTxBody processes a txBody element, returning one string per paragraph.
func extractTxBody(d *xml.Decoder) []string {
	var lines []string
	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "p" {
				text := extractParagraphText(d)
				if text != "" {
					lines = append(lines, text)
				}
			} else {
				d.Skip()
			}
		case xml.EndElement:
			return lines
		}
	}
	return lines
}

// extractParagraphText processes a paragraph, concatenating all run text.
func extractParagraphText(d *xml.Decoder) string {
	var sb strings.Builder
	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "r" {
				sb.WriteString(extractRunText(d))
			} else {
				d.Skip()
			}
		case xml.EndElement:
			return sb.String()
		}
	}
	return sb.String()
}

// extractRunText processes a run element, returning its text content.
func extractRunText(d *xml.Decoder) string {
	var text string
	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "t" {
				var s string
				d.DecodeElement(&s, &t)
				text += s
			} else {
				d.Skip()
			}
		case xml.EndElement:
			return text
		}
	}
	return text
}

func readPptx(path string) ([]string, error) {
	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("unable to open file: %w", err)
	}
	defer r.Close()

	// Collect slide files and sort numerically
	type slideEntry struct {
		num  int
		file *zip.File
	}
	var slides []slideEntry
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			name := strings.TrimPrefix(f.Name, "ppt/slides/slide")
			name = strings.TrimSuffix(name, ".xml")
			num, err := strconv.Atoi(name)
			if err != nil {
				continue
			}
			slides = append(slides, slideEntry{num: num, file: f})
		}
	}
	sort.Slice(slides, func(i, j int) bool {
		return slides[i].num < slides[j].num
	})

	var lines []string
	for i, s := range slides {
		if i > 0 {
			lines = append(lines, "")
		}
		lines = append(lines, fmt.Sprintf("--- Slide %d ---", s.num))

		data, err := readFileFromZip(r, s.file.Name)
		if err != nil {
			return nil, fmt.Errorf("error reading slide %d: %w", s.num, err)
		}

		d := xml.NewDecoder(strings.NewReader(string(data)))
		lines = append(lines, extractTextFromXML(d)...)
	}

	return lines, nil
}
