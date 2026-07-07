"""Shared filesystem path configuration."""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv


PROJECT_ROOT = Path(__file__).resolve().parents[1]
load_dotenv(PROJECT_ROOT / ".env")

DB_MIS_DIR = Path(r"E:\DB_MIS")

def sampled_db_path(filename: str, env_name: str | None = None) -> Path:
    """Resolve an external workbook path from .env or DB_MIS_DIR."""
    if env_name:
        raw = os.getenv(env_name)
        if raw:
            return Path(raw).expanduser()
    return DB_MIS_DIR / filename


def sampled_db_path_str(filename: str, env_name: str | None = None) -> str:
    """String form for legacy scripts and argparse defaults."""
    return str(sampled_db_path(filename, env_name))
