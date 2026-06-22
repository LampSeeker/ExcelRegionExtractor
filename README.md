# Excel Region Extractor

Excel workbook의 각 sheet에서 정보가 들어 있는 영역을 찾아 Excel range 문자열로 저장하는 Python 프로젝트입니다.

값이 있는 셀, 병합 셀, border 구조, 삽입 이미지를 이용해 정보영역 bbox를 추출하고 JSON, overlay PNG, embedded image 파일을 생성합니다.

## 설치

```powershell
pip install -e .
```

## 실행

전체 sheet 실행:

```powershell
excel-regions --workbook examples/sample.xlsx --out outputs/all_sheets
```

특정 sheet만 실행:

```powershell
excel-regions --workbook examples/sample.xlsx --sheet "각형맨홀(특2호)" --out outputs/manhole
```

overlay PNG 없이 JSON만 생성:

```powershell
excel-regions --workbook examples/sample.xlsx --out outputs/all_sheets --no-images
```

호환 wrapper:

```powershell
python scripts/extract_info_regions.py --workbook examples/sample.xlsx --out outputs/all_sheets
```

Python 코드에서 사용:

```python
from excel_info_region import extract_workbook_info_regions
from excel_info_region.config import load_config

config = load_config("config/default.json")
result = extract_workbook_info_regions("examples/sample.xlsx", config=config)
```

## 출력

```text
outputs/all_sheets/
  info_regions_full.json
  info_regions_summary.json

  각형맨홀(특2호)/
    info_regions.json
    info_regions.png
    images/
      IMG001_B4_H13_Picture_1.png
```

sheet별 `info_regions.json`:

```json
{
  "sheet_name": "각형맨홀(특2호)",
  "regions": [
    "A1:P1",
    "B4:H13",
    "H4:P12"
  ],
  "images": [
    {
      "name": "Picture 1",
      "range_ref": "B4:H13",
      "path": "images/IMG001_B4_H13_Picture_1.png"
    }
  ]
}
```

`regions`는 정보영역 range 문자열 목록입니다. `images`는 추출된 embedded image의 이름, anchor range, 상대 경로만 기록합니다.

## 처리 흐름

```text
Excel workbook
  -> 값이 있는 셀 수집
  -> 값이 있는 병합 셀은 병합 범위 전체를 occupied 처리
  -> occupied cell connected component 추출
  -> border/table shell 기반 bbox 보정
  -> border contact merge로 인접 정보영역 병합
  -> embedded image anchor를 별도 image region/link로 기록
  -> sheet별 JSON, workbook summary JSON, 선택적 overlay PNG 저장
```

이미지는 cell occupied와 섞지 않습니다. 도면 이미지와 바로 아래 표가 붙어 있어도 하나의 정보영역으로 과병합하지 않기 위한 의도입니다.

## 설정

기본 설정은 `config/default.json`입니다.

주요 옵션:

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

overlay PNG에서 한글이 깨지면 font path를 지정합니다.

```json
{
  "visualization": {
    "font_path": "C:/Windows/Fonts/malgun.ttf"
  }
}
```

`--no-images`는 overlay PNG 생성을 끄는 옵션입니다. embedded image 추출은 `extract_embedded_images` 설정이 `true`이면 계속 수행됩니다.

## 프로젝트 구조

```text
src/excel_info_region/
  extractor.py       workbook/sheet 실행 흐름과 최종 region 조합
  cells.py           값/병합 셀 occupied 처리
  components.py      connected component, bbox 중복 제거, 기하 helper
  borders.py         border shell 보정, border-contact merge
  image_regions.py   sheet image anchor를 region box로 변환
  image_export.py    embedded image 파일 추출
  raw_drawing.py     xlsx DrawingML 직접 파싱
  runner.py          CLI 실행, JSON/PNG/image 저장
  visualize.py       debug overlay PNG 렌더링
```

## 개발

테스트:

```powershell
pytest
```

최소 회귀 확인:

```powershell
pytest
excel-regions --workbook examples/sample.xlsx --out outputs/all_sheets --no-images
```

시각화나 image 추출을 바꿨다면 `--no-images` 없이 실행해서 sheet별 `info_regions.png`와 `images/` 결과를 확인합니다.

## 주의

`openpyxl`은 수식을 직접 계산하지 않습니다. overlay PNG는 `data_only=True` workbook을 사용하므로 Excel 파일 안에 cached formula result가 없으면 수식 결과가 비어 보일 수 있습니다. 필요한 경우 Excel에서 파일을 한 번 열고 저장한 뒤 실행합니다.
