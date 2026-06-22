# Agent Guide

## Summary

`excel-region-extractor` extracts information-region ranges from Excel workbooks.

It reads cell values, merged cells, borders, and embedded image anchors, then writes:

- sheet JSON: `regions` and `images`
- workbook summary JSON
- optional overlay PNG
- extracted embedded image files

## Commands

Install locally:

```powershell
pip install -e .
```

Run the public sample:

```powershell
excel-regions --workbook examples/synthetic_demo.xlsx --out outputs/demo
```

Run without overlay PNG:

```powershell
excel-regions --workbook examples/synthetic_demo.xlsx --out outputs/demo --no-images
```

Test:

```powershell
pytest
```

## Important Files

```text
src/excel_info_region/
  cli.py             console entrypoint
  runner.py          writes JSON, overlay PNG, extracted images
  extractor.py       workbook/sheet orchestration
  cells.py           cell and merged-cell occupied logic
  borders.py         border expansion and border-contact merge
  components.py      connected components and bbox helpers
  image_regions.py   image anchors to region boxes
  image_export.py    embedded image extraction
  raw_drawing.py     raw xlsx DrawingML parsing
  visualize.py       overlay PNG renderer
```

Other paths:

```text
config/default.json
examples/synthetic_demo.xlsx
docs/images/synthetic_demo_regions.png
tests/
```

## Output Schema

Sheet JSON:

```json
{
  "sheet_name": "Synthetic Demo",
  "regions": ["A1:H1", "A3:D6"],
  "images": [
    {
      "name": "Picture 1",
      "range_ref": "G4:I9",
      "path": "images/IMG001_G4_I9_Picture_1.png"
    }
  ]
}
```

Internal extractor results may still use `info_regions`; `runner.py` converts the user-facing JSON key to `regions`.

## Rules

- Do not commit private Excel samples. `examples/sample.xlsx` and `examples/sample2.xlsx` are ignored local files.
- Keep the public sample synthetic and non-sensitive.
- Keep images separate from cell connected components to avoid over-merging drawings and nearby tables.
- `openpyxl` does not calculate formulas. Overlay rendering uses `data_only=True`, so formula results require cached values saved by Excel.
- If output schema changes, update README and tests in the same change.
