# office-reader

Command-line tool to extract text content from Microsoft Office files (`.docx`, `.pptx`, `.xlsx`), using only the Go standard library.

Office Open XML files are ZIP archives containing XML. This program decompresses them and parses the XML to extract text.

## Installation

```bash
go build -o office-reader
```

## Usage

```bash
./office-reader <file.docx|.pptx|.xlsx>
```

Each output line is prefixed with its index:

```
[0] First line
[1] Second line
```

### Supported formats

| Format | Extracted content |
|--------|-------------------|
| `.docx` | Paragraphs, tables, hyperlinks (Markdown format `[text](url)`) |
| `.pptx` | Shape text, separated by slide (`--- Slide N ---`) |
| `.xlsx` | Cells formatted as `\| col1 \| col2 \|`, separated by sheet (`--- SheetName ---`) |

## Project structure

```
main.go   — entry point, dispatch by extension
zip.go    — shared ZIP/XML helpers
docx.go   — .docx reader
pptx.go   — .pptx reader
xlsx.go   — .xlsx reader
```

## Dependencies

None — only the Go standard library (`archive/zip`, `encoding/xml`).
