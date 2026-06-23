from __future__ import annotations

import json
from importlib.resources import files
from pathlib import Path
from typing import Any


def load_config(path: str | Path | None = None) -> dict[str, Any]:
    if path is None:
        return json.loads(files("excel_info_region").joinpath("default.json").read_text(encoding="utf-8"))
    path = Path(path)
    return json.loads(path.read_text(encoding="utf-8"))
