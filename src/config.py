"""
File: config.py
Project: excel-ps-batch-export
Description: Environment and path configuration for code/data separation.
"""

from __future__ import annotations

import os
import sys


def load_dotenv() -> None:
    """Load .env from the project root without overriding existing environment."""
    dotenv_path = os.path.join(PROJECT_DIR, ".env")
    if not os.path.exists(dotenv_path):
        return
    with open(dotenv_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key, value = key.strip(), value.strip()
            if (value.startswith('"') and value.endswith('"')) or (
                value.startswith("'") and value.endswith("'")
            ):
                value = value[1:-1]
            if key and key not in os.environ:
                os.environ[key] = value


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)

load_dotenv()

# EPS_DATA_DIR separates code from data.
# If set, workspace / export / log.csv live under that path.
# If not set, the project uses demo/ as the data directory.
DATA_DIR = os.environ.get("EPS_DATA_DIR", os.path.join(PROJECT_DIR, "demo"))
WORKSPACE_DIR = os.path.join(DATA_DIR, "workspace")
EXPORT_DIR = os.path.join(DATA_DIR, "export")
LOG_PATH = os.path.join(DATA_DIR, "log.csv")
FONTS_DIR = os.path.join(WORKSPACE_DIR, "assets", "fonts")
FONTS_CONFIG_PATH = os.path.join(WORKSPACE_DIR, "fonts.json")
DEFAULT_FONT_PATH = os.path.join(FONTS_DIR, "AlibabaPuHuiTi-2-85-Bold.ttf")

# One module object whether imported as `config` or `src.config`.
sys.modules["config"] = sys.modules[__name__]
sys.modules["src.config"] = sys.modules[__name__]
