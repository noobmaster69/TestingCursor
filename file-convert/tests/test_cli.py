import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_list_formats():
    result = subprocess.run(
        [sys.executable, "-m", "file_convert", "--list-formats"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=ROOT,
    )
    assert result.returncode == 0, result.stderr
    assert "md" in result.stdout and "html" in result.stdout


def test_convert_md_html(fixtures_dir: Path, tmp_path: Path):
    src = fixtures_dir / "sample.md"
    out = tmp_path / "report.html"
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "file_convert",
            str(src),
            "--to",
            "html",
            "-o",
            str(out),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=ROOT,
    )
    assert result.returncode == 0, result.stderr + result.stdout
    assert out.is_file()
