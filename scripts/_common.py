from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from excel_region_lab.config import load_config
from excel_region_lab.io import ensure_dir, write_json
from excel_region_lab.signals import extract_sheet_signals, iter_target_sheets, open_workbook
from excel_region_lab.components import connected_components


def parser(description: str) -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description=description)
    p.add_argument("--workbook", default=str(PROJECT_ROOT / "examples" / "sample.xlsx"))
    p.add_argument("--sheet", default=None)
    p.add_argument("--config", default=str(PROJECT_ROOT / "config" / "default.json"))
    p.add_argument("--out", default=str(PROJECT_ROOT / "outputs"))
    return p


def load_context(args):
    config = load_config(args.config)
    wb = open_workbook(args.workbook)
    out = ensure_dir(args.out)
    return config, wb, out


def sheet_contexts(args):
    config, wb, out = load_context(args)
    for ws in iter_target_sheets(wb, args.sheet):
        signals = extract_sheet_signals(ws, config["signals"])
        primitives = connected_components(signals, config["connected_components"])
        yield config, out, ws, signals, primitives
