package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"sort"
	"strconv"
	"strings"
)

type SlideContent struct {
	CSld struct {
		SpTree struct {
			Shapes []Shape `xml:"sp"`
		} `xml:"spTree"`
	} `xml:"cSld"`
}

type Shape struct {
	TxBody *TxBody `xml:"txBody"`
}

type TxBody struct {
	Paragraphs []AParagraph `xml:"p"`
}

type AParagraph struct {
	Runs []ARun `xml:"r"`
}

type ARun struct {
	Text string `xml:"t"`
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

		var slide SlideContent
		if err := xml.Unmarshal(data, &slide); err != nil {
			return nil, fmt.Errorf("error parsing slide %d: %w", s.num, err)
		}

		for _, shape := range slide.CSld.SpTree.Shapes {
			if shape.TxBody == nil {
				continue
			}
			for _, p := range shape.TxBody.Paragraphs {
				var sb strings.Builder
				for _, run := range p.Runs {
					sb.WriteString(run.Text)
				}
				text := sb.String()
				if text != "" {
					lines = append(lines, text)
				}
			}
		}
	}

	return lines, nil
}
