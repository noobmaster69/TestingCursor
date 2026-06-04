# Universal File Converter — Product & Technical Spec

**Project codename:** `file-convert`  
**Status:** Planning  
**Parent doc:** [python-productivity-ideas.md](./python-productivity-ideas.md) (idea #1)

---

## 1. Summary

A local-first CLI that converts one file at a time between common formats—documents, spreadsheets, images, PDFs, and lightweight archives—without uploading data to the cloud. The tool is designed around a **plugin registry**: each conversion pair is a small handler; the CLI resolves the route, validates options, and writes output beside the source (or to a user-specified path).

**One-liner:** `convert <input> --to <format>` replaces bookmarked web converters and manual “Save As” loops.

---

## 2. Goals and non-goals

### Goals (v1)

- Single command, predictable output path, clear errors.
- Pure-Python paths where possible (MD, CSV, XLSX); subprocess only when necessary (LibreOffice, optional Ghostscript).
- Extensible registry so adding `webp → png` is one new module + registration line.
- Safe defaults: dry-run for external binaries; never overwrite input without explicit `-o` targeting a different file.
- Works on Windows (primary), macOS, and Linux with documented optional dependencies.

### Non-goals (v1)

| Deferred | Reason |
|----------|--------|
| Batch folder conversion | Adds queueing, progress, cancellation—ship single-file first |
| Watch folder / daemon | Operational complexity |
| OCR (scanned PDF → text) | Different product surface |
| Video/audio transcoding | Heavy deps, licensing |
| Cloud APIs | Privacy and offline requirement |
| DRM-protected Office | Legal/technical edge cases |
| GUI | CLI proves value; GUI later |

---

## 3. Users and use cases

| Persona | Job to be done |
|---------|----------------|
| Developer | `report.md --to html` for quick preview; `openapi.json` N/A in v1 |
| Analyst | `export.xlsx --to csv --sheet "Q1"` before pandas ingest |
| Designer / PM | `deck.pdf` page 3 → PNG for Slack; `photo.heic --to jpg` |
| Student / office worker | `essay.docx --to pdf` before submission |

**Example session**

```bash
# Preview what LibreOffice would run
convert essay.docx --to pdf --dry-run

# Convert spreadsheet tab
convert data.xlsx --to csv --sheet "Q1 Summary" -o ./out/

# PDF page range to images
convert scan.pdf --to png --pages 1-3 --dpi 200 -o ./pages/

# List zip contents without extracting
convert bundle.zip --list
```

---

## 4. CLI specification

### 4.1 Command shape

```
convert INPUT_PATH --to FORMAT [OPTIONS]
convert INPUT_PATH --list              # zip only (alias of --to list)
convert --list-formats                 # show supported (from, to) pairs
convert --doctor                       # check optional deps (LibreOffice, etc.)
```

### 4.2 Global options

| Flag | Type | Default | Description |
|------|------|---------|-------------|
| `--to` | string | required* | Target format extension without dot (`pdf`, `csv`, `html`) |
| `-o`, `--output` | path | auto | File or directory; see [Output rules](#7-output-path-rules) |
| `--from-format` | string | inferred | Override source format when extension is wrong/missing |
| `--dry-run` | bool | false | Print planned actions; no writes |
| `--force` | bool | false | Overwrite existing output |
| `--verbose` | bool | false | Full tracebacks, subprocess stdout |
| `--timeout` | int | 120 | Subprocess cap (seconds) for LibreOffice etc. |

\*Not required when using `--list` on zip or `convert --list-formats`.

### 4.3 Per-domain options

**Spreadsheets (xlsx ↔ csv)**

| Flag | Description |
|------|-------------|
| `--sheet` | Sheet name or 0-based index; default first sheet |
| `--encoding` | CSV output encoding (default `utf-8`) |
| `--delimiter` | CSV delimiter (default `,`) |

**Markdown → HTML**

| Flag | Description |
|------|-------------|
| `--theme` | `github`, `minimal`, or path to `.css` |
| `--title` | HTML `<title>`; default stem of input file |

**PDF ↔ images**

| Flag | Description |
|------|-------------|
| `--pages` | `1`, `1-5`, `1,3,7-9`; default all pages |
| `--dpi` | Rasterization DPI (default `150`) |
| `--quality` | JPEG quality 1–100 (default `85`) |

**Images (incl. HEIC)**

| Flag | Description |
|------|-------------|
| `--quality` | Output JPEG/WebP quality |
| `--max-width` / `--max-height` | Resize preserving aspect if set |

**Office (docx → pdf, etc.)**

| Flag | Description |
|------|-------------|
| `--dry-run` | **Recommended default in docs** for first-time users |

**Zip**

| Flag | Description |
|------|-------------|
| `--list` | Table of names, sizes, compressed sizes (no extraction) |

### 4.4 Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | User error (unsupported pair, bad flags, missing file) |
| 2 | Missing system dependency (LibreOffice not found) |
| 3 | Conversion failed (corrupt file, subprocess error) |
| 130 | Interrupted (Ctrl+C) |

### 4.5 User-facing messages

- **Unsupported pair:** `Cannot convert .docx → .png. Supported: docx→pdf, pdf→png, ... Run: convert --list-formats`
- **Missing dep:** `LibreOffice not found. Install: https://... or run: convert --doctor`

---

## 5. Supported conversions (v1 matrix)

| From | To | Engine | v1 priority |
|------|-----|--------|-------------|
| `.md` | `.html` | `markdown` + Pygments | P0 |
| `.csv` | `.xlsx` | `pandas` / `openpyxl` | P0 |
| `.xlsx` | `.csv` | `openpyxl` read_only | P0 |
| `.pdf` | `.png` / `.jpg` | `pymupdf` | P1 |
| `.png`, `.jpg`, `.webp` | `.pdf` | Pillow + `img2pdf` | P1 |
| `.heic` | `.jpg` | `pillow-heif` | P1 (optional extra) |
| `.docx`, `.odt` | `.pdf` | LibreOffice headless | P2 |
| `.zip` | — | list only (`--list`) | P1 |

**P0** = milestone 1; **P1** = milestone 2; **P2** = milestone 3.

Future v2 pairs (not in v1): `html→pdf` (weasyprint), `json→csv`, multi-image → zip, `pptx→pdf`.

---

## 6. Architecture

### 6.1 Repository layout

```
file-convert/
├── pyproject.toml
├── README.md
├── src/
│   └── file_convert/
│       ├── __init__.py
│       ├── __main__.py          # python -m file_convert
│       ├── cli.py               # Typer or argparse entry
│       ├── doctor.py            # dependency probes
│       ├── models.py            # ConversionJob, ConversionResult
│       ├── registry.py          # register + resolve handlers
│       ├── output.py            # path resolution, collision policy
│       ├── errors.py            # typed exceptions → exit codes
│       └── handlers/
│           ├── __init__.py      # auto-register all handlers
│           ├── base.py          # BaseHandler protocol
│           ├── markdown_html.py
│           ├── spreadsheet.py
│           ├── pdf_images.py
│           ├── images_pdf.py
│           ├── heic_jpeg.py
│           ├── office_pdf.py
│           └── zip_list.py
├── tests/
│   ├── fixtures/                # tiny md, csv, xlsx, pdf, zip
│   ├── test_registry.py
│   ├── test_output_paths.py
│   └── handlers/
│       └── test_markdown_html.py
└── docs/
    └── universal-file-converter.md   # this file (or link from monorepo)
```

### 6.2 Core types

```python
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol

@dataclass(frozen=True)
class ConversionJob:
    source: Path
    source_format: str          # normalized: "md", "xlsx"
    target_format: str
    output: Path | None         # None → resolver picks
    options: dict[str, Any] = field(default_factory=dict)
    dry_run: bool = False
    force: bool = False

@dataclass(frozen=True)
class ConversionResult:
    output_path: Path
    duration_ms: int
    message: str | None = None  # e.g. "3 pages written"

class Handler(Protocol):
    """One module may implement multiple (from, to) pairs."""

    @property
    def pairs(self) -> frozenset[tuple[str, str]]: ...

    def validate(self, job: ConversionJob) -> None:
        """Raise ValidationError before any IO."""

    def convert(self, job: ConversionJob) -> ConversionResult:
        """Perform conversion or dry-run preview."""
```

### 6.3 Registry

```python
class Registry:
    def register(self, handler: Handler) -> None: ...
    def resolve(self, src_fmt: str, tgt_fmt: str) -> Handler: ...
    def list_pairs(self) -> list[tuple[str, str]]: ...
```

**Resolution flow**

```
CLI parse args
    → infer source_format from extension or --from-format
    → Registry.resolve(src, tgt)
    → handler.validate(job)
    → output.resolve_output_path(job)  # may raise if collision
    → handler.convert(job)
    → print result path (rich green check optional)
```

**Registration** (at import time):

```python
# handlers/__init__.py
def build_registry() -> Registry:
    r = Registry()
    for handler in [MarkdownHtml(), SpreadsheetHandler(), ...]:
        r.register(handler)
    return r
```

### 6.4 Format normalization

- Lowercase, strip leading dot: `.PDF` → `pdf`.
- Aliases: `jpeg` → `jpg`, `tif` → `tiff`, `markdown` → `md`.
- Unknown extension: require `--from-format` or fail with hint.

---

## 7. Output path rules

| `-o` value | Behavior |
|------------|----------|
| omitted | Same directory as source; stem + new extension: `report.md` → `report.html` |
| file path | Exact file; create parent dirs if needed |
| directory | `dir / {stem}.{target}` |
| PDF → multi-page PNG | `dir / {stem}_page_{n:04d}.png` |

**Collision policy**

- If output exists and not `--force`: error with message suggesting `-o` or `--force`.
- Never use input path as output path for in-place transforms (v1 has none).

---

## 8. Handler implementation notes

### 8.1 Markdown → HTML (P0)

**Libraries:** `markdown`, `pygments` (fenced code), optional `jinja2` for wrapper template.

**Steps**

1. Read UTF-8 text; optional BOM strip.
2. `markdown.markdown(text, extensions=["fenced_code", "tables", "toc"])`.
3. Wrap in HTML5 skeleton; inject CSS from `--theme` (bundled assets in package `data/themes/`).
4. Write `output.html`.

**Dry-run:** Log source size, target path, theme name.

**Tests:** Golden file compare (normalize whitespace); broken UTF-8 → exit 3.

---

### 8.2 XLSX ↔ CSV (P0)

**XLSX → CSV**

- `openpyxl.load_workbook(..., read_only=True, data_only=True)`.
- Resolve sheet by name or index; error if missing with list of sheet names.
- Stream rows to csv.writer; no full DataFrame required for large files.

**CSV → XLSX**

- `pandas.read_csv` with delimiter/encoding from options (reasonable for v1).
- `to_excel` with single sheet name from `--sheet` or `"Sheet1"`.
- Optional: column width auto-fit (stretch)—nice-to-have.

**Risks**

| Issue | Mitigation |
|-------|------------|
| 500MB xlsx RAM spike | `read_only=True`; document `--sample` in v2 |
| Dates as numbers | `data_only=True`; document Excel date limitation |
| Formulas | Export computed values only |

---

### 8.3 PDF → images (P1)

**Library:** `pymupdf` (fitz).

**Steps**

1. Open PDF; parse `--pages` into sorted unique 0-based indices.
2. Per page: `page.get_pixmap(dpi=dpi)` → PNG or JPEG per target format.
3. Multi-page: multiple files with `_page_NNNN` suffix unless single page → single file.

**Progress:** `rich.progress` when page count > 5.

**Dry-run:** Print page list and estimated file count.

---

### 8.4 Images → PDF (P1)

**Libraries:** Pillow for normalize; `img2pdf` for assembly.

- Single image → single-page PDF.
- v2: multiple inputs `convert a.jpg b.jpg --to pdf` (out of v1 scope).

**Color modes:** Convert RGBA to RGB for JPEG-backed PDF if needed.

---

### 8.5 HEIC → JPG (P1, optional extra)

**Extra:** `pip install file-convert[heic]` → `pillow-heif` register opener.

- Fail at handler registration if import missing; `--doctor` reports status.
- Honor `--quality`; preserve EXIF orientation if trivial with Pillow.

**Windows note:** README section on wheel availability and VC++ runtime if needed.

---

### 8.6 Office → PDF (P2)

**Engine:** LibreOffice subprocess.

```text
soffice --headless --convert-to pdf --outdir <tmpdir> <input.docx>
```

**Implementation**

- `shutil.which("soffice")` or Windows `soffice.exe` common paths.
- Run in temp dir; move result to resolved output path.
- `--timeout` kill tree on hang.
- **Dry-run:** print exact command and binary path.

**Supported inputs (v1):** `.docx`, `.odt` (LibreOffice native). Not `.doc` unless LO supports on platform—document limitation.

---

### 8.7 Zip list (P1)

- Stdlib `zipfile`; no extraction.
- Output table: `name | size | compress_size | is_dir`.
- `--to list` or dedicated `convert file.zip --list`.

---

## 9. Dependency and packaging

### 9.1 `pyproject.toml` sketch

```toml
[project]
name = "file-convert"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = [
  "typer>=0.12",
  "rich>=13",
  "markdown>=3.5",
  "pygments>=2.17",
  "openpyxl>=3.1",
  "pandas>=2.2",
  "pymupdf>=1.24",
  "Pillow>=10",
  "img2pdf>=0.5",
]

[project.optional-dependencies]
heic = ["pillow-heif>=0.16"]
dev = ["pytest>=8", "ruff>=0.4"]

[project.scripts]
convert = "file_convert.cli:app"
```

### 9.2 Optional system binaries

| Tool | Used for | Detection |
|------|----------|-----------|
| LibreOffice `soffice` | docx→pdf | `convert --doctor` |
| Ghostscript | future compress | not v1 |

---

## 10. `convert --doctor` output (example)

```text
file-convert 0.1.0

Core (Python)
  markdown ............... OK
  pymupdf ................ OK
  openpyxl ............... OK

Optional
  pillow-heif ............ MISSING  (pip install file-convert[heic])
  LibreOffice ............ OK  C:\Program Files\LibreOffice\program\soffice.exe

Ready: 6/8 conversion pairs available
```

---

## 11. Security and safety

- **No eval** on user content; parsers only.
- **Zip slip:** N/A for v1 (no extract); v2 extract must validate paths.
- **Subprocess:** no shell=True; argument list form; temp dir cleaned in `finally`.
- **Path traversal:** resolve `input` and `-o` to absolute paths; reject `..` in output if outside intended dir (optional strict mode later).
- **Secrets:** files stay local; no network calls in core.

---

## 12. Testing strategy

| Layer | What to test |
|-------|----------------|
| Registry | duplicate registration fails; resolve unknown pair |
| Output resolver | directory vs file; collision without `--force` |
| Handlers | golden outputs on fixtures; dry-run produces no files |
| Office | mock subprocess; skip integration CI without LO |
| CLI | `subprocess.run(["convert", ...])` e2e on tiny fixtures |

**Fixtures (minimal)**

- `sample.md`, `sample.csv`, `sample.xlsx` (2 sheets), `sample.pdf` (2 pages), `sample.zip`, `sample.heic` (optional LFS or skip CI).

**CI:** GitHub Actions matrix: ubuntu + windows-latest; `pytest -m "not office"` default.

---

## 13. Implementation milestones

### Milestone 1 — Core + P0 (week 1)

- [ ] Project scaffold, Typer CLI, Registry, models, errors
- [ ] Output path resolver + tests
- [ ] MD→HTML with one bundled theme
- [ ] XLSX↔CSV both directions
- [ ] `convert --list-formats`
- [ ] README quickstart

**Definition of done:** `convert tests/fixtures/sample.md --to html` works offline on Windows.

### Milestone 2 — PDF, images, zip (week 2)

- [ ] PDF→PNG with `--pages` and `--dpi`
- [ ] PNG/JPG→PDF
- [ ] Zip `--list`
- [ ] HEIC handler behind optional extra
- [ ] Progress bar for multi-page PDF

### Milestone 3 — Office + polish (week 3)

- [ ] LibreOffice wrapper + dry-run + timeout
- [ ] `convert --doctor`
- [ ] Exit code documentation
- [ ] Example GIF in README (optional)

---

## 14. Open questions

| # | Question | Proposal |
|---|----------|----------|
| 1 | Typer vs argparse? | Typer for ergonomics; keep handlers free of CLI |
| 2 | Monorepo vs standalone repo? | Standalone `file-convert` repo when implementation starts |
| 3 | Batch in v1.1? | `convert ./dir/*.md --to html` with glob and `--jobs 4` |
| 4 | Config file? | Defer; flags sufficient for v1 |

---

## 15. README outline (for implementation)

1. Install: `pip install file-convert` / optional `[heic]`
2. Three copy-paste examples (md→html, xlsx→csv, pdf→png)
3. Table of supported pairs linking to `--list-formats`
4. LibreOffice install links (Win/Mac/Linux)
5. Troubleshooting: SSL not relevant; HEIC missing; LibreOffice PATH

---

## 16. Link back to productivity roadmap

This spec elaborates **idea #1** from [python-productivity-ideas.md](./python-productivity-ideas.md). When v1 ships, update that file with a link: `See [universal-file-converter.md](./universal-file-converter.md) for full spec.`

**Suggested next step:** Milestone 1 scaffold—`registry.py` + `markdown_html.py` handler + three pytest files—before any Office or HEIC work.
