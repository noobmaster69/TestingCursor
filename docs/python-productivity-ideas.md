# Python Productivity Tools — Planning Notes

Planning doc for small Python utilities that save time on everyday tasks: converting files, batch edits, organizing folders, and glue work between apps. Written for a solo developer or small team building reliable CLI-first tools.

---

## 1. Universal File Converter (CLI)

**Implementation:** [file-convert/](../file-convert/) — see [universal-file-converter.md](./universal-file-converter.md) for full spec.

### Concept

One command to convert between common formats locally—no cloud upload, no account.

```bash
convert input.docx --to pdf
convert data.xlsx --to csv --sheet "Q1 Summary"
convert photo.heic --to jpg --quality 85
convert archive.zip --list
convert report.md --to html --theme github
```

### Why it’s cool

Replaces a dozen bookmarked websites and ad-hoc “open in Excel then Save As” workflows. Great portfolio piece if the plugin architecture is clean.

### MVP scope (v1)

| In scope | Out of scope (v1) |
|----------|-------------------|
| Single input file → single output | Watch folder / hot folder daemon |
| PDF ↔ images (page to PNG) | OCR / scanned PDF → text |
| CSV ↔ XLSX | Video/audio transcoding |
| MD → HTML | DRM-protected Office files |
| HEIC → JPG (if lib available) | Batch zip of 1000+ files without progress |

### Architecture

```
file-convert/
├── src/
│   ├── cli.py
│   ├── registry.py      # format pair → handler
│   └── handlers/
│       ├── pdf_images.py
│       ├── office_csv.py
│       └── markdown_html.py
├── tests/fixtures/
└── pyproject.toml
```

```python
@dataclass
class ConversionJob:
    source: Path
    target_format: str
    options: dict[str, Any]

class Converter(Protocol):
    supported: set[tuple[str, str]]  # ("docx", "pdf")
    def convert(self, job: ConversionJob) -> Path: ...
```

### Tech stack (by format)

| Conversion | Library / tool | Notes |
|------------|----------------|-------|
| MD → HTML | `markdown` + Pygments | Optional CSS from `--theme` |
| XLSX → CSV | `openpyxl` or `pandas` | `--sheet` name or index |
| CSV → XLSX | `pandas` | Preserve column widths optional |
| PDF → PNG | `pymupdf` (fitz) | `--dpi 150` |
| Images → PDF | `Pillow` + `img2pdf` | Multi-image merge |
| DOCX → PDF | `libreoffice --headless` subprocess | Document dependency in README |
| HEIC → JPG | `pillow-heif` | Windows install notes |

### UX details

- **Dry-run:** `convert file.docx --to pdf --dry-run` prints command that would run (especially for LibreOffice).
- **Output path:** default `input.pdf` beside source; override with `-o out/`.
- **Errors:** exit code 1 + one-line reason; `--verbose` for stack trace.

### Milestones

1. Registry + MD→HTML + CSV↔XLSX (pure Python).
2. PDF/image handlers.
3. Office via LibreOffice wrapper with timeout.
4. Progress bar (`rich`) for multi-page PDF export.

### Risks

| Risk | Mitigation |
|------|------------|
| LibreOffice not installed | Detect at startup; link to install docs |
| Huge XLSX memory | `read_only=True` in openpyxl; chunk CSV writes |
| Wrong MIME guess | Trust extension + optional `--from-format` |

---

## 2. Bulk Rename + Organize by Rules

### Concept

Productivity sibling to a “smart renamer”: sort downloads into folders by extension, date, or keyword; optional duplicate detection before move.

```bash
organize ~/Downloads --preset downloads
organize ./photos --by date-taken --format YYYY/MM
organize . --rule 'invoice*.pdf' --move ./Finance/Invoices
```

### MVP

- Presets: `downloads`, `documents`, `images` (yaml-defined extension → folder map)
- `--dry-run` default; `--execute` to apply
- Log file: `organize-2026-06-04.jsonl` for audit

### Rule engine (yaml)

```yaml
rules:
  - match: "*.pdf"
    move_to: Documents/PDF
  - match: "*report*"
    move_to: Work/Reports
  - match: "*.png"
    move_to: Images/Screenshots
    if_modified_within_days: 7
```

### Implementation

- `pathlib` + `shutil.move`; handle name collisions with `(1)` suffix policy.
- Duplicate detection: SHA-256 first 1MB + file size quick hash; full hash on `--deep`.

---

## 3. Clipboard History & Snippet Manager

### Concept

Lightweight tray-less CLI: save clipboard entries, search snippets, paste back to clipboard (Windows/macOS/Linux).

```bash
clip save --tag api
clip list --limit 20
clip search "Bearer"
clip paste 3          # copy item #3 to system clipboard
clip snippet add deploy -f deploy.sh
```

### MVP

- SQLite: `id`, `content`, `content_type` (text/url/path), `tags`, `created_at`
- Max entries configurable; prune oldest
- **Windows:** `pyperclip` or `win32clipboard` for text; skip images in v1

### Privacy

- Local DB only; `--exclude-pattern` for credit cards (regex redact on save)
- Optional `--lock` password for DB (stretch)

---

## 4. PDF Toolkit (Merge, Split, Compress, Watermark)

### Concept

Daily PDF chores without opening Acrobat.

```bash
pdf merge a.pdf b.pdf -o combined.pdf
pdf split combined.pdf --pages 1-3,7 -o part.pdf
pdf compress big.pdf --target-size 5mb
pdf watermark draft.pdf --text "DRAFT" --opacity 0.3
```

### Stack

- `pymupdf` for merge/split/watermark
- Compress: rasterize images downscale vs. `ghostscript` subprocess for aggressive shrink

### Safety

- Never overwrite input without `-o`; validate page ranges before write

---

## 5. Text / Data Format Transformer

### Concept

Pipe or file-based conversions developers hit constantly.

```bash
xfmt input.json --to yaml
xfmt data.csv --to json --records column
xfmt config.ini --to env
xfmt timestamps.log --to iso --column 0
xfmt secrets.env --to json --mask
```

### MVP pairs

- JSON ↔ YAML ↔ TOML (use `ruamel.yaml` for round-trip friendly YAML)
- CSV → JSON lines
- `.env` ↔ dict export (no eval; line-based parser)
- Unix timestamp column → ISO8601 in CSV

### Design

- stdin support: `cat file | xfmt --to yaml`
- `--mask` replaces values matching `KEY_*=`, `password`, etc.

---

## 6. Screenshot & File “Inbox” Processor

### Concept

Watch a folder (or run on demand) that normalizes screenshots: OCR optional, rename with timestamp, move to project folder based on frontmost window title (OS-specific) or filename pattern.

```bash
inbox process ~/Screenshots --rename "{date}_{time}_screenshot"
inbox process . --ocr --lang en  # optional tesseract
```

### MVP

- On-demand only (no daemon): `process` scans folder once
- Rename template: `{date}`, `{time}`, `{original_stem}`
- Move rules from #2 shared yaml `rules` module (reuse library)

### Stretch

- `watchdog` for real-time folder watch
- Tesseract for searchable PNG sidecar `.txt`

---

## 7. Email / Report Attachment Extractor

### Concept

Pull attachments from `.eml`, `.msg` (if lib), or mbox export into structured folders.

```bash
attachments extract mail.mbox --out ./attachments --types pdf,xlsx
attachments list mail.eml
```

### MVP

- `.eml` via stdlib `email` package
- Filter by extension; dedupe by hash
- Summary CSV: `filename, subject, date, size`

---

## 8. Scheduled “Folder Sync” Mirror (One-way backup)

### Concept

Mirror `Documents/Projects` → `Backup/Projects` with delete propagation optional—simpler than rsync for Windows users.

```bash
sync run projects.yaml
sync projects.yaml --dry-run
```

```yaml
# projects.yaml
jobs:
  - name: projects-backup
    source: C:/Users/me/Projects
    dest: D:/Backup/Projects
    exclude: ["node_modules", ".git", "__pycache__"]
    delete_extra: false
```

### Implementation

- Compare mtime + size (fast); `--checksum` for paranoia
- `shutil.copy2`; log actions to JSONL
- Task Scheduler / cron example in README

---

## 9. Quick Launcher + Command Aliases (dotfiles helper)

### Concept

YAML-defined shortcuts that expand to shell commands or open paths/apps—portable across machines.

```bash
run dev          # → cd project && code .
run standup      # → python git_digest.py --since 7d
run convert-pdf  # → convert $1 --to pdf
```

### MVP

- `run <alias>` with `$1` positional substitution
- `run --list` shows aliases from `~/.config/runner/config.yaml`
- Cross-platform: invoke via `subprocess` shell=True on Windows carefully (document injection risk; no arbitrary shell from untrusted yaml)

---

## 10. Meeting Notes → Action Items Parser

### Concept

Paste or read a `.txt` / `.md` meeting note; extract lines that look like tasks (TODO, `- [ ]`, names + verbs) into `tasks.md` or export to Todoist/JSON (stub).

```bash
tasks extract notes.md --out tasks.md
tasks extract notes.md --format json
```

### MVP

- Heuristics: lines starting with `- [ ]`, `TODO:`, `Action:`, `@name`
- Optional LLM pass behind `--smart` flag (env API key)
- No external API required for v1

---

## 11. Image Batch Processor (Resize, Strip EXIF, WebP)

### Concept

Photography and web prep in one CLI.

```bash
img batch ./raw --resize 1920x1080 --keep-aspect --out ./web
img batch ./raw --format webp --quality 80
img batch ./raw --strip-exif
```

### Stack

- `Pillow` pipeline; preserve folder structure with `--mirror-tree`
- Parallel: `ProcessPoolExecutor` for CPU-bound encodes

---

## 12. Password / Secret Generator & Vault Lite

### Concept

Generate passphrases, API keys, `.env` entries; store in encrypted local vault (not a full 1Password competitor).

```bash
secret generate --length 32 --charset alnum
secret generate passphrase --words 5
secret set api.stripe --value sk_live_...
secret get api.stripe --copy
```

### MVP

- Generation only + append to `.env.example` style file
- Vault v2: `cryptography` Fernet with key derived from master password (document threat model)

---

## Comparison — what to build first

| Idea | Daily utility | Complexity | Best dependency lesson |
|------|---------------|------------|---------------------------|
| File Converter | Very high | Medium | Plugin registry, subprocess |
| PDF Toolkit | High | Low–medium | pymupdf |
| Format Transformer | High (dev) | Low | stdin/stdout CLI |
| Bulk Organize | High | Medium | Rule engine + safety |
| Image Batch | Medium | Low | Pillow, parallelism |
| Clipboard Manager | Medium | Medium | OS APIs, SQLite |
| Folder Sync | Medium | Medium | diff strategies |
| Attachment Extractor | Low–medium | Low | email stdlib |
| Quick Launcher | Medium | Low | YAML config |
| Task Parser | Medium | Low (+ optional AI) | text heuristics |

**Suggested first build:** **Format Transformer (`xfmt`)** or **PDF Toolkit**—small scope, pure Python, immediate daily use. **File Converter** is the flagship if you want one umbrella project with plugins.

---

## Shared product principles

1. **CLI first** — scriptable, composable with pipes; GUI optional later.
2. **Dry-run by default** for anything that moves, renames, or deletes files.
3. **Local-first** — no telemetry; secrets stay in env vars.
4. **Explicit dependencies** — optional extras in `pyproject.toml`:
   - `pip install file-convert[pdf,office]`
5. **Exit codes** — `0` success, `1` user error, `2` system/dependency missing.
6. **Windows + macOS + Linux** — test path handling and LibreOffice/Ghostscript discovery on each.

---

## Monorepo option

If several of these ship together:

```
productivity-tools/
├── packages/
│   ├── convert/      # idea 1
│   ├── xfmt/         # idea 5
│   ├── pdftool/      # idea 4
│   └── organize/     # idea 2
├── shared/
│   └── ptools_common/  # logging, dry-run, collision policy
└── README.md
```

Shared module avoids copy-paste for: colored logging (`rich`), dry-run wrapper, config yaml loader, and `ensure_dir(path)`.

---

## Next actions

1. Pick one tool and copy its **MVP scope** into a GitHub issue.
2. Spike the hardest dependency (e.g. DOCX→PDF on your OS).
3. Ship v0.1 with `--help` and three README examples; iterate from real friction in your own Downloads folder.
