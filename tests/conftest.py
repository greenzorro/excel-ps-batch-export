"""Pytest configuration: unit tests use demo/ as DATA_DIR by default.

Forces EPS_DATA_DIR to demo/ for the suite. temporary_data_dir() still
overrides paths per test when needed.
"""

import os
from pathlib import Path

_DEMO_DIR = Path(__file__).resolve().parent.parent / "demo"
os.environ["EPS_DATA_DIR"] = str(_DEMO_DIR)
