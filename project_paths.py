"""
Hardcoded project paths for Google Colab and local runs.

Colab — run the pipeline:
    from google.colab import drive
    drive.mount("/content/drive")
    exec(open("/content/Siemens-Air/run_pipeline.py").read())

Layout:
  - Python code:  /content/Siemens-Air/
  - Data folders: Google Drive (input / processing / output)
"""

from __future__ import annotations

from pathlib import Path

# Where .py files live (upload scripts here in Colab).
_COLAB_CODE_DIRS = (
    Path("/content/Siemens-Air"),
    Path("/content/Siemens-air"),
)


def _resolve_script_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        pass
    for path in _COLAB_CODE_DIRS:
        if path.is_dir():
            return path
    return Path.cwd()


_SCRIPT_DIR = _resolve_script_dir()

BASE_DIR = next((path for path in _COLAB_CODE_DIRS if path.is_dir()), _SCRIPT_DIR)

# Google Drive data root (input / processing / output).
_DRIVE_DATA_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT Siemens/Siemens Air"
)

if _DRIVE_DATA_ROOT.is_dir():
    INPUT_DIR = _DRIVE_DATA_ROOT / "input"
    PROCESSING_DIR = _DRIVE_DATA_ROOT / "processing"
    OUTPUT_DIR = _DRIVE_DATA_ROOT / "output"
else:
    INPUT_DIR = BASE_DIR / "input"
    PROCESSING_DIR = BASE_DIR / "processing"
    OUTPUT_DIR = BASE_DIR / "output"


def ensure_project_dirs() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
