# Excel Region Extractor

Extract Excel information-region ranges from workbook sheets.

The tool uses cell values, merged cells, borders, and embedded image anchors to write JSON outputs, optional overlay PNGs, and extracted embedded image files.

## Install

From PyPI:

```powershell
pip install excel-region-extractor
```

From GitHub:

```powershell
pip install git+https://github.com/LampSeeker/ExcelRegionExtractor.git
```

For local development:

```powershell
pip install -e .
```

## Usage

Run all sheets:

```powershell
excel-regions --workbook examples/synthetic_demo.xlsx --out outputs/all_sheets
```

Run one sheet:

```powershell
excel-regions --workbook examples/synthetic_demo.xlsx --sheet "Synthetic Demo" --out outputs/demo
```

Write JSON and extracted embedded images without overlay PNG:

```powershell
excel-regions --workbook examples/synthetic_demo.xlsx --out outputs/all_sheets --no-images
```

Python API:

```python
from excel_info_region import extract_workbook_info_regions
from excel_info_region.config import load_config

config = load_config("config/default.json")
result = extract_workbook_info_regions("examples/synthetic_demo.xlsx", config=config)
```

## Output

```text
outputs/all_sheets/
  info_regions_full.json
  info_regions_summary.json

  Synthetic Demo/
    info_regions.json
    info_regions.png
    images/
      IMG001_G4_I9_Image_1.png
```

Sheet JSON:

```json
{
  "sheet_name": "Synthetic Demo",
  "regions": [
    "A1:H1",
    "A3:D6",
    "G4:I9",
    "A9:E12"
  ],
  "images": [
    {
      "name": "Image 1",
      "range_ref": "G4:I9",
      "path": "images/IMG001_G4_I9_Image_1.png"
    }
  ]
}
```

`regions` is the list of detected Excel ranges. `images` records extracted embedded image metadata and relative file paths.

Example overlay:

![Synthetic Excel region overlay](docs/images/synthetic_demo_regions.png)

## Processing Flow

```text
Excel workbook
  -> collect non-empty cells
  -> expand non-empty merged cells to their full merged ranges
  -> find occupied-cell connected components
  -> expand ranges with border/table shells
  -> merge adjacent regions by border contact
  -> keep embedded image regions separate
  -> write sheet JSON, workbook summary JSON, and optional overlay PNG
```

Images are intentionally kept separate from cell connected components. This avoids over-merging drawings with nearby tables.

## Configuration

Default config:

```text
config/default.json
```

Common options:

```json
{
  "include_values": true,
  "include_merged_cells": true,
  "include_images": true,
  "include_grouped_drawing_images": true,
  "use_borders": true,
  "strong_borders_only": true,
  "use_border_contact_merge": true,
  "extract_embedded_images": true,
  "embedded_image_dir": "images"
}
```

Set a font path if Korean text is broken in overlay PNGs:

```json
{
  "visualization": {
    "font_path": "C:/Windows/Fonts/malgun.ttf"
  }
}
```

`--no-images` skips overlay PNG generation. Embedded image extraction still runs when `extract_embedded_images` is `true`.

## Project Structure

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

## Development

```powershell
pytest
excel-regions --workbook examples/synthetic_demo.xlsx --out outputs/all_sheets --no-images
```

Run without `--no-images` when changing visualization or image extraction.

## Notes

`openpyxl` does not calculate formulas. Overlay rendering uses `data_only=True`, so formula cells need cached values saved by Excel to show calculated results.
