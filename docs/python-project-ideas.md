# Python Project Ideas — Planning Notes

Developer planning doc for small-to-medium Python scripts that are fun to build, useful to keep, and good portfolio pieces. Each idea includes intent, scope, and enough technical detail to start implementation.

---

## 1. Local Data Pipeline Inspector

### Concept

A CLI tool that scans a folder (CSV, JSON, Parquet, SQLite), profiles each dataset, and writes a single HTML or Markdown report: row counts, null rates, dtypes, duplicate keys, and simple outliers.

### Why it’s cool

Turns “what’s in this folder?” from manual `pandas` snippets into a repeatable audit you can run before notebooks or ETL jobs.

### Users

- You, before exploratory analysis
- Teammates onboarding to messy exports

### MVP scope (v1)

| In scope | Out of scope (later) |
|----------|----------------------|
| Recursive scan of one root path | Scheduled runs / daemon |
| CSV + JSON + SQLite | Cloud storage (S3, GCS) |
| Per-file summary stats | ML-based anomaly detection |
| Combined report output | Web UI |

### Technical approach

```
pipeline-inspector/
├── src/
│   ├── cli.py           # argparse / typer entry
│   ├── scanners/        # one adapter per format
│   ├── profile.py       # shared stats (nulls, uniques, min/max)
│   └── report.py        # Jinja2 → HTML or MD template
├── tests/
└── pyproject.toml
```

- **CLI:** `typer` or `argparse`; flags: `--path`, `--output report.html`, `--formats csv,json`.
- **Profiling:** `pandas` for tabular; `sqlite3` + `PRAGMA` for DB files; stream large CSVs with `chunksize` to avoid RAM spikes.
- **Report:** Jinja2 template listing files as cards; link to per-file detail section.
- **Config:** optional `pipeline-inspector.yaml` for ignore globs (`__pycache__`, `*.tmp`).

### Data model (internal)

```python
@dataclass
class FileProfile:
    path: Path
    format: str
    row_count: int | None
    columns: list[ColumnStats]
    warnings: list[str]  # e.g. "duplicate id", "90% null in email"
```

### Milestones

1. **Day 1–2:** CSV scanner + console table output.
2. **Day 3–4:** HTML report + JSON/ SQLite adapters.
3. **Day 5:** Tests with fixture files; README with example screenshots.

### Risks and mitigations

| Risk | Mitigation |
|------|------------|
| Huge files OOM | Chunked reads; cap rows profiled with `--sample 10000` |
| Slow on many files | Parallel scan with `concurrent.futures` (IO-bound) |
| False “duplicate key” on sampled data | Document sampling; full scan flag `--full` |

### Success criteria

- Profiles 50 mixed files in under 30s on a laptop.
- Report is readable without opening Python.

---

## 2. Git Change Digest (Daily Standup Helper)

### Concept

Script that runs in a repo (or list of repos), collects commits and diff stats since a date, and generates a bullet summary for standup or weekly notes—grouped by author, directory, or conventional commit type.

### Why it’s cool

Automates the boring “what did I ship this week?” step; teaches `subprocess` + `git` plumbing without a heavy framework.

### MVP scope

- Single repo path; `--since 7d` or `--since 2026-06-01`
- Output: Markdown file or stdout
- Group by: author (default), optional `--by path` (top-level folders)

### Technical approach

- Use `git log --pretty=format:%H|%an|%s|%ad` and `git diff --stat` per commit or aggregate with `git log --stat`.
- Parse conventional commits (`feat:`, `fix:`) with regex; bucket “other”.
- Optional: `anthropic` / local LLM pass to turn raw subjects into 3 narrative bullets (feature-flagged, API key from env).

### CLI sketch

```bash
python git_digest.py --repo . --since 7d --out standup.md
```

### Dependencies

- Stdlib only for v1 (`subprocess`, `pathlib`, `datetime`)
- Optional: `python-dateutil` for friendly `--since` parsing

### Tests

- Fixture repo with known commits (init temp dir, scripted commits in `pytest` fixture).

### Stretch goals

- Multi-repo config file (`repos: [path1, path2]`)
- GitHub PR titles via `gh pr list` for the same date range

---

## 3. Smart File Renamer (Rules Engine)

### Concept

Batch-rename files from patterns: dates in EXIF, regex capture groups, sequential padding, and dry-run preview before apply.

### Why it’s cool

Everyone has `IMG_2847.jpg` chaos; a safe dry-run + undo log is a real utility.

### MVP scope

- One directory, non-recursive (v1)
- Rule types: `regex`, `replace`, `prefix`, `sequential`
- `--dry-run` (default) vs `--execute`
- Write `rename.log.json` for rollback script

### Architecture

```python
class RenameRule(Protocol):
    def apply(self, path: Path) -> Path: ...

class Pipeline:
    rules: list[RenameRule]
    def preview(self, files: list[Path]) -> list[tuple[Path, Path]]: ...
```

- **Safety:** refuse to overwrite unless `--force`; cap collisions.
- **Undo:** `undo.py` reads log and swaps names back.

### Example rules file (`rules.yaml`)

```yaml
directory: ./photos
rules:
  - type: regex
    pattern: '^IMG_(\d+)\.jpg$'
    replacement: 'vacation_2026_\1.jpg'
  - type: sequential
    start: 1
    width: 4
```

### Tech choices

- `PyYAML` for rules; `rich` for terminal preview table.
- Optional: `Pillow` for EXIF-based `date_taken` rename rule.

### Edge cases to plan for

- Case-insensitive filesystems (Windows)
- Open files / permission errors — collect failures, don’t partial-apply without transaction semantics (apply in two phases: all valid renames planned, then execute)

---

## 4. API Contract Mock Server

### Concept

Given an OpenAPI 3 YAML/JSON spec, spin up a local Flask/FastAPI server that returns example responses and validates incoming requests against the schema—useful for frontend or mobile dev when the real backend isn’t ready.

### Why it’s cool

Bridges design-first workflows; one spec drives docs, mocks, and later real implementation checks.

### MVP scope

- Load one OpenAPI file
- Implement `GET` routes with static examples from `example` or `examples` in spec
- Return `501` for unimplemented methods with clear JSON body
- Log request bodies that fail validation

### Stack

- **FastAPI** + `openapi-core` or **prance** + **jsonschema** for validation
- **Faker** to generate data when examples are missing (seed for reproducibility)

### Project layout

```
mock-server/
├── main.py
├── loader.py      # parse & normalize spec
├── handlers.py    # dynamic route registration
└── fixtures/      # optional override JSON per path
```

### Milestones

1. Parse spec and list routes.
2. Register routes with example responses.
3. Request validation middleware.
4. CLI: `mock-server serve api.yaml --port 8080`

### Risks

- OpenAPI `$ref` resolution — use a dedicated resolver library, don’t hand-roll.
- Auth endpoints — stub `401/403` with configurable behavior.

---

## 5. Personal Metrics CLI (Quantified Self Lite)

### Concept

Log habits or measurements from the terminal (`water 250`, `run 5km`, `mood 4`), store in SQLite, and plot weekly trends or export CSV.

### Why it’s cool

Minimal friction capture; you own the data; good practice for schema design and small visualizations.

### Schema (v1)

```sql
CREATE TABLE events (
  id INTEGER PRIMARY KEY,
  metric TEXT NOT NULL,      -- 'water_ml', 'run_km', 'mood'
  value REAL NOT NULL,
  note TEXT,
  recorded_at TEXT NOT NULL  -- ISO8601 UTC
);
CREATE INDEX idx_metric_time ON events(metric, recorded_at);
```

### Commands

```bash
metrics log water 250
metrics log mood 4 --note "afternoon slump"
metrics show water --days 7
metrics export --format csv --out events.csv
```

### Implementation notes

- **Storage:** `sqlite3` stdlib; migration script when schema changes.
- **Plots:** `matplotlib` optional extra `[plots]` so core stays lightweight.
- **Validation:** predefined metrics in `metrics.yaml` with unit and min/max.

### Privacy

- Local-only by default; no network calls in v1.

---

## Picking one to build first

| Idea | Build time | Learning value | Daily utility |
|------|------------|----------------|---------------|
| Pipeline Inspector | ~1 week | pandas, reporting | High if you work with data |
| Git Digest | 2–3 days | git, CLI design | Medium |
| File Renamer | 3–5 days | safety, UX | High once |
| Mock Server | ~1–2 weeks | APIs, OpenAPI | High for full-stack teams |
| Metrics CLI | 3–4 days | SQLite, viz | Personal |

**Recommendation:** Start with **Git Change Digest** or **Pipeline Inspector**—both fit this repo’s data/Python focus, ship a useful v1 quickly, and leave clear extension paths.

---

## Shared engineering standards (all projects)

- **Python:** 3.11+; type hints on public functions.
- **Packaging:** `pyproject.toml` + `uv` or `pip`; one console script entry point per project.
- **Testing:** `pytest`; target critical paths (parsers, rename preview, git parsing), not 100% coverage for v1.
- **Docs:** README with install, one example command, and limitations.
- **CI (optional):** GitHub Action running `ruff check` + `pytest` on push.

---

## Next step

Choose one idea, open a tracking issue (or `docs/ROADMAP.md` section), and define v1 acceptance criteria before writing feature code. Spike the riskiest part first (e.g. chunked CSV profiling or OpenAPI `$ref` loading).
