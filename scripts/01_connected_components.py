from __future__ import annotations

from _common import parser, sheet_contexts, write_json


def main() -> None:
    args = parser("Test Connected Component Labeling").parse_args()
    for _, out, ws, signals, primitives in sheet_contexts(args):
        data = {
            "sheet_name": ws.title,
            "bounds": signals.bounds.to_dict() if signals.bounds else None,
            "algorithm": "connected_components",
            "regions": [r.to_dict() for r in primitives],
        }
        path = write_json(out / f"{ws.title}_connected_components_regions.json", data)
        print(f"[connected_components] {ws.title}: {len(primitives)} regions -> {path}")


if __name__ == "__main__":
    main()
