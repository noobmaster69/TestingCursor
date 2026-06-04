# file-convert

Local-first CLI to convert files between common formats—no cloud upload.

```bash
cd file-convert
pip install -e .

# Prefer python -m or `file-convert` on Windows (avoids CONVERT.EXE name clash)
python -m file_convert sample.md --to html --theme github
file-convert data.xlsx --to csv --sheet "Sheet1"
file-convert scan.pdf --to png --pages 1-3 --dpi 200
file-convert bundle.zip --list
file-convert --doctor
```

## Supported conversions (v0.1)

| From | To | Notes |
|------|-----|--------|
| md | html | Themes: `github`, `minimal`, or path to `.css` |
| xlsx | csv | `--sheet` name or index |
| csv | xlsx | `--sheet` for output tab name |
| pdf | png, jpg | `--pages`, `--dpi` |
| png, jpg, webp, tiff | pdf | Single image |
| heic | jpg | Requires `pip install -e ".[heic]"` |
| docx, odt | pdf | Requires [LibreOffice](https://www.libreoffice.org/) |
| zip | list | `convert file.zip --list` |

Run `convert --list-formats` for the live list on your machine.

## Options

```
convert INPUT --to FORMAT [-o OUTPUT] [--dry-run] [--force] [--verbose]
```

Format-specific: `--sheet`, `--encoding`, `--delimiter`, `--theme`, `--title`, `--pages`, `--dpi`, `--quality`, `--max-width`, `--max-height`

## Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | User error |
| 2 | Missing dependency |
| 3 | Conversion failed |

## Development

```bash
pip install -e ".[dev]"
pytest
```

## Spec

See [../docs/universal-file-converter.md](../docs/universal-file-converter.md) for the full product and technical specification.
