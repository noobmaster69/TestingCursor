# file-convert

Convert files on your computer—no upload. Use the **GUI** (easiest) or the **command line**.

## Quick start (GUI)

```bash
cd file-convert
pip install -e .

# Double-click (Windows) or run:
launch_gui.bat

# Or from terminal:
python -m file_convert.gui
file-convert-gui
file-convert --gui
```

### Using the window

1. **Browse** — pick your file (Word, Excel, PDF, images, markdown, zip, etc.).
2. **Convert to** — choose from formats that work for that file.
3. **Save location** — optional; leave blank to save beside the original.
4. Click **Convert** — then **Open output folder** or **Open output file**.

Options appear when relevant (Excel sheet name, HTML theme, PDF pages/DPI).

## Command line

```bash
python -m file_convert sample.md --to html --theme github
file-convert data.xlsx --to csv --sheet "Sheet1"
file-convert scan.pdf --to png --pages 1-3 --dpi 200
file-convert bundle.zip --list
file-convert --doctor
```

On Windows, prefer `python -m file_convert` or `file-convert` — the bare `convert` command can clash with `CONVERT.EXE`.

## Supported conversions

| From | To | Notes |
|------|-----|--------|
| md | html | Themes: `github`, `minimal` |
| xlsx | csv | Sheet name in GUI / `--sheet` |
| csv | xlsx | |
| pdf | png, jpg | Pages & DPI in GUI |
| png, jpg, webp, tiff | pdf | |
| heic | jpg | `pip install -e ".[heic]"` |
| docx, odt | pdf | [LibreOffice](https://www.libreoffice.org/) |
| zip | list | View contents in status log |

## Development

```bash
pip install -e ".[dev]"
pytest
```

## Spec

[../docs/universal-file-converter.md](../docs/universal-file-converter.md)
