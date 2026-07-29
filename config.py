"""Shared filesystem path configuration."""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv


# .env 는 이 파일과 같은 폴더(레포 루트)에 있다. 상위 프로젝트에서 분리되기 전에는
# 한 단계 위였으므로, 이전 위치도 함께 시도해 둔다(먼저 로드된 값이 우선).
PROJECT_ROOT = Path(__file__).resolve().parent
load_dotenv(PROJECT_ROOT / ".env")
load_dotenv(PROJECT_ROOT.parent / ".env")

# 수집 원본/산출물 공유 폴더 — 웹앱(BEMS)의 SAMPLED_DB_DIR 과 동일해야 한다.
DB_MIS_DIR = Path(os.getenv("SAMPLED_DB_DIR", r"E:\DB_MIS")).expanduser()

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
