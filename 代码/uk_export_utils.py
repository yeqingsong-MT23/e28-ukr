# -*- coding: utf-8 -*-
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path


INVALID_FOLDER_CHARS_RE = re.compile(r'[\\/:*?"<>|]+')


def _build_time_range_output_dir(export_dir: Path, base_dir: Path | None) -> Path:
    if base_dir is None:
        return export_dir

    time_range_file = base_dir / "时间范围.txt"
    if not time_range_file.exists():
        return export_dir

    values = [
        line.strip().lstrip("\ufeff")
        for line in time_range_file.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    if len(values) < 2:
        return export_dir

    folder_name = f"{values[0]}-{values[1]}"
    folder_name = INVALID_FOLDER_CHARS_RE.sub("-", folder_name).strip().rstrip(".")
    if not folder_name:
        return export_dir

    return export_dir / folder_name


def build_default_output_path(export_dir: Path, default_output_name: str, base_dir: Path | None = None) -> Path:
    target_dir = _build_time_range_output_dir(export_dir, base_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    timestamp = dt.datetime.now().strftime("%m%d%H%M")
    file_name = f"{Path(default_output_name).stem}-{timestamp}{Path(default_output_name).suffix}"
    return target_dir / file_name
