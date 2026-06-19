from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def load_config(path: str | Path | None = None) -> dict[str, Any]:
    if path is None:
        path = Path(__file__).resolve().parents[2] / "config" / "default.json"
    path = Path(path)
    return json.loads(path.read_text(encoding="utf-8"))
