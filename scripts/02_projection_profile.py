from __future__ import annotations

from _common import parser, sheet_contexts, write_json
from excel_region_lab.profiles import projection_profile


def main() -> None:
    args = parser("Test Projection Profile").parse_args()
    for config, out, ws, signals, _ in sheet_contexts(args):
        data = projection_profile(signals, config["projection_profile"])
        path = write_json(out / f"{ws.title}_projection_profiles.json", data)
        print(
            f"[projection_profile] {ws.title}: "
            f"blank_rows={len(data['blank_row_runs'])}, blank_cols={len(data['blank_col_runs'])} -> {path}"
        )


if __name__ == "__main__":
    main()
