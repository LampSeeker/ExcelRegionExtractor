# Excel Info Region Extractor

엑셀 시트에서 **정보영역 bbox만** 추출하는 최소 프로젝트입니다.

이 프로젝트는 다음을 하지 않습니다.

```text
role_hint 없음
semantic relation 없음
layout root 없음
structural tag 없음
title/body 판단 없음
```

출력은 오직 `info_regions`입니다.

## 처리 구조

```text
Excel workbook
  ↓
값이 있는 셀 수집
  ↓
값이 있는 병합 셀은 병합 범위 전체를 occupied 처리
  ↓
cell occupied만 connected component로 묶음
  ↓
삽입 이미지는 별도 image region으로 추가
  ↓
cell region과 image region을 서로 병합하지 않음
  ↓
info_regions.json / info_regions.png 출력
```

이미지를 cell occupied와 섞지 않기 때문에, 도면 이미지와 바로 아래 표가 붙어 있어도 하나로 과병합하지 않습니다.

## 설치

```powershell
cd excel_info_region_extractor
pip install -e .
```

## 전체 시트 실행

```powershell
python scripts/run_all.py --workbook examples/sample.xlsx --out outputs/all_sheets
```

## 특정 시트 실행

```powershell
python scripts/extract_info_regions.py --workbook examples/sample.xlsx --sheet "각형맨홀(특2호)" --out outputs/manhole
```

## 이미지 출력 없이 JSON만 실행

```powershell
python scripts/run_all.py --workbook examples/sample.xlsx --out outputs/all_sheets --no-images
```

## 출력 구조

```text
outputs/all_sheets/
  info_regions_full.json
  info_regions_summary.json

  각형맨홀(특2호)/
    info_regions.json
    info_regions.png

  U형측구/
    info_regions.json
    info_regions.png
```

`info_regions.json` 예시:

```json
{
  "sheet_name": "각형맨홀(특2호)",
  "info_regions": [
    "A1:P1",
    "B4:H13",
    "H4:P12",
    "R4:S9",
    "R11:S11",
    "E13:L22",
    "C23:N30",
    "A32:P60"
  ]
}
```

## 한글 폰트

PNG에서 한글이 깨지면 `config/default.json`의 `visualization.font_path`를 설정하세요.

Windows:

```json
{
  "visualization": {
    "font_path": "C:/Windows/Fonts/malgun.ttf"
  }
}
```

WSL:

```json
{
  "visualization": {
    "font_path": "/mnt/c/Windows/Fonts/malgun.ttf"
  }
}
```

## 수식 결과값

PNG 렌더링은 `data_only=True`로 workbook을 열기 때문에, 엑셀 파일 안에 cached formula result가 있으면 수식 문자열 대신 결과값이 보입니다.

단, `openpyxl`은 수식을 직접 계산하지 않습니다. 결과값이 비어 있으면 Excel에서 파일을 한 번 열고 저장해야 합니다.


## v2 변경 사항

`info_regions`를 객체 목록이 아니라 range 문자열 목록으로 단순화했습니다.

```json
{
  "sheet_name": "각형맨홀(특2호)",
  "info_regions": ["A1:P1", "B4:H13"]
}
```
