package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"strings"
)

type Document struct {
	Body Body `xml:"body"`
}

type BodyItem struct {
	Paragraph *Paragraph
	Table     *Table
}

type Body struct {
	Items []BodyItem
}

func (b *Body) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "p":
				var p Paragraph
				if err := d.DecodeElement(&p, &t); err != nil {
					return err
				}
				b.Items = append(b.Items, BodyItem{Paragraph: &p})
			case "tbl":
				var table Table
				if err := d.DecodeElement(&table, &t); err != nil {
					return err
				}
				b.Items = append(b.Items, BodyItem{Table: &table})
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

type Table struct {
	Rows []TableRow `xml:"tr"`
}

type TableRow struct {
	Cells []TableCell `xml:"tc"`
}

type TableCell struct {
	Paragraphs []Paragraph `xml:"p"`
}

type Run struct {
	Text Text `xml:"t"`
}

type Text struct {
	Value string `xml:",chardata"`
}

type Hyperlink struct {
	Id   string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	Runs []Run  `xml:"r"`
}

type ParagraphItem struct {
	Run       *Run
	Hyperlink *Hyperlink
}

type Paragraph struct {
	Items []ParagraphItem
}

func (p *Paragraph) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "r":
				var r Run
				if err := d.DecodeElement(&r, &t); err != nil {
					return err
				}
				p.Items = append(p.Items, ParagraphItem{Run: &r})
			case "hyperlink":
				var h Hyperlink
				if err := d.DecodeElement(&h, &t); err != nil {
					return err
				}
				p.Items = append(p.Items, ParagraphItem{Hyperlink: &h})
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

func renderParagraph(p *Paragraph, rels map[string]string) string {
	var sb strings.Builder
	for _, item := range p.Items {
		if item.Run != nil {
			sb.WriteString(item.Run.Text.Value)
		} else if item.Hyperlink != nil {
			var linkText strings.Builder
			for _, r := range item.Hyperlink.Runs {
				linkText.WriteString(r.Text.Value)
			}
			url := rels[item.Hyperlink.Id]
			if url != "" {
				sb.WriteString(fmt.Sprintf("[%s](%s)", linkText.String(), url))
			} else {
				sb.WriteString(linkText.String())
			}
		}
	}
	return sb.String()
}

func readDocx(path string) ([]string, error) {
	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("unable to open file: %w", err)
	}
	defer r.Close()

	rels, err := readRelsFromZip(r, "word/_rels/document.xml.rels", "hyperlink")
	if err != nil {
		return nil, err
	}

	data, err := readFileFromZip(r, "word/document.xml")
	if err != nil {
		return nil, err
	}
	if data == nil {
		return nil, fmt.Errorf("document.xml not found")
	}

	var doc Document
	if err := xml.Unmarshal(data, &doc); err != nil {
		return nil, fmt.Errorf("XML parsing error: %w", err)
	}

	var lines []string
	for _, bodyItem := range doc.Body.Items {
		if bodyItem.Paragraph != nil {
			lines = append(lines, renderParagraph(bodyItem.Paragraph, rels))
		} else if bodyItem.Table != nil {
			for _, row := range bodyItem.Table.Rows {
				var cells []string
				for _, cell := range row.Cells {
					var cellParts []string
					for _, p := range cell.Paragraphs {
						cellParts = append(cellParts, renderParagraph(&p, rels))
					}
					cells = append(cells, strings.Join(cellParts, " "))
				}
				lines = append(lines, "| "+strings.Join(cells, " | ")+" |")
			}
		}
	}

	return lines, nil
}
