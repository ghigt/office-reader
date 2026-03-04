package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

type SharedStrings struct {
	Items []SharedStringItem `xml:"si"`
}

type SharedStringItem struct {
	Text string  `xml:"t"`
	Runs []SRun  `xml:"r"`
}

type SRun struct {
	Text string `xml:"t"`
}

type Workbook struct {
	Sheets struct {
		Items []WBSheet `xml:"sheet"`
	} `xml:"sheets"`
}

type WBSheet struct {
	Name    string `xml:"name,attr"`
	SheetId string `xml:"sheetId,attr"`
}

type Worksheet struct {
	SheetData SheetData `xml:"sheetData"`
}

type SheetData struct {
	Rows []XRow `xml:"row"`
}

type XRow struct {
	Cells []XCell `xml:"c"`
}

type XCell struct {
	Type  string `xml:"t,attr"`
	Value string `xml:"v"`
}

func readXlsx(path string) ([]string, error) {
	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("unable to open file: %w", err)
	}
	defer r.Close()

	// 1. Parse shared strings
	var sst []string
	sstData, err := readFileFromZip(r, "xl/sharedStrings.xml")
	if err != nil {
		return nil, err
	}
	if sstData != nil {
		var ss SharedStrings
		if err := xml.Unmarshal(sstData, &ss); err != nil {
			return nil, fmt.Errorf("sharedStrings parsing error: %w", err)
		}
		for _, item := range ss.Items {
			if item.Text != "" {
				sst = append(sst, item.Text)
			} else {
				var sb strings.Builder
				for _, run := range item.Runs {
					sb.WriteString(run.Text)
				}
				sst = append(sst, sb.String())
			}
		}
	}

	// 2. Read sheet names from workbook
	wbData, err := readFileFromZip(r, "xl/workbook.xml")
	if err != nil {
		return nil, err
	}
	if wbData == nil {
		return nil, fmt.Errorf("workbook.xml not found")
	}
	var wb Workbook
	if err := xml.Unmarshal(wbData, &wb); err != nil {
		return nil, fmt.Errorf("workbook parsing error: %w", err)
	}

	// 3. Read each sheet
	var lines []string
	for i, sheet := range wb.Sheets.Items {
		if i > 0 {
			lines = append(lines, "")
		}
		lines = append(lines, fmt.Sprintf("--- %s ---", sheet.Name))

		sheetPath := fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)
		data, err := readFileFromZip(r, sheetPath)
		if err != nil {
			return nil, fmt.Errorf("error reading %s: %w", sheetPath, err)
		}
		if data == nil {
			continue
		}

		var ws Worksheet
		if err := xml.Unmarshal(data, &ws); err != nil {
			return nil, fmt.Errorf("error parsing %s: %w", sheetPath, err)
		}

		for _, row := range ws.SheetData.Rows {
			var cells []string
			for _, cell := range row.Cells {
				value := cell.Value
				if cell.Type == "s" {
					idx, err := strconv.Atoi(value)
					if err == nil && idx >= 0 && idx < len(sst) {
						value = sst[idx]
					}
				}
				cells = append(cells, value)
			}
			lines = append(lines, "| "+strings.Join(cells, " | ")+" |")
		}
	}

	return lines, nil
}
