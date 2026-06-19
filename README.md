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


## v3 변경 사항: border-aware 보정

출력은 여전히 range 문자열 목록만 유지합니다.

```json
{
  "sheet_name": "토공",
  "info_regions": [
    "B2:G2",
    "B4:G8"
  ]
}
```

내부 알고리즘에만 border/table shell 보정을 추가했습니다.

```text
1. 값/병합셀 기반 occupied 생성
2. 값/병합셀 connected component 추출
3. border가 있는 셀을 별도 occupied로 생성
4. border connected component 추출
5. value component가 border shell 안에 있거나 충분히 겹치면 bbox를 border shell까지 확장
6. 이미지 anchor region은 여전히 별도로 추가
7. 최종 출력은 range 문자열만 반환
```

설정은 `config/default.json`에서 조정합니다.

```json
{
  "use_borders": true,
  "strong_borders_only": true,
  "min_border_cells": 2,
  "border_expand_min_value_overlap": 0.8,
  "border_expand_min_border_overlap": 0.1,
  "border_expand_max_area_ratio": 3.0,
  "border_expand_max_extra_rows": 3,
  "border_expand_max_extra_cols": 3,
  "add_border_only_regions": false
}
```

`add_border_only_regions`는 기본 `false`입니다.  
즉 값이 전혀 없는 테두리 박스만으로는 최종 정보영역을 만들지 않습니다.


기본값은 보수적으로 설정했습니다. 큰 border shell이 제목/도면/다른 표까지 삼키지 않도록 `border_expand_max_extra_rows`, `border_expand_max_area_ratio` 제한을 둡니다.


## v4 변경 사항: border contact merge

출력은 계속 range 문자열 목록만 유지합니다.

v4는 최후 보정으로 `border contact merge`를 추가했습니다.

```text
1. 기존 info region 후보를 먼저 만든다.
2. 셀 border를 edge 그래프로 만든다.
3. edge endpoint가 서로 닿는 border edge들을 하나의 border component로 묶는다.
4. 각 info region의 외곽선이 어떤 border component와 실제로 접촉하는지 확인한다.
5. 두 개 이상의 info region이 같은 border component에 접촉하고, 서로 가까운 수직/수평 이웃이면 병합한다.
6. 단순히 같은 border bbox 안에 들어간다는 이유만으로는 병합하지 않는다.
7. image region은 이 병합 단계에 섞지 않는다.
```

관련 설정:

```json
{
  "use_border_contact_merge": true,
  "border_contact_strong_only": false,
  "border_contact_tolerance_cells": 0,
  "border_contact_min_edges": 1,
  "border_contact_merge_max_gap": 1,
  "border_contact_merge_min_axis_overlap": 0.8,
  "border_contact_merge_max_area_ratio": 2.5
}
```


## v5 변경 사항: 2-side border contact merge

v5는 v4의 border contact merge를 더 일반화했습니다.

기존 v4:

```text
같은 border component에 접촉
+ gap <= 1
+ 축 overlap 충분함
```

v5:

```text
같은 border component에 접촉
+ 각 region이 같은 border component의 최소 2개 변 이상에 접촉
+ gap <= 5
+ 축 overlap 충분함
+ 병합 bbox가 과도하게 커지지 않음
```

이 규칙은 특정 시트명, 특정 좌표, 특정 문구를 사용하지 않습니다.  
셀 border의 기하 구조만 사용합니다.

관련 설정:

```json
{
  "use_border_contact_merge": true,
  "border_contact_min_touched_sides": 2,
  "border_contact_min_edges_per_side": 1,
  "border_contact_merge_max_gap": 5,
  "border_contact_merge_min_axis_overlap": 0.8,
  "border_contact_merge_max_area_ratio": 2.5
}
```

출력 형식은 계속 range 문자열 목록입니다.

```json
{
  "sheet_name": "토공",
  "info_regions": [
    "B16:J53"
  ]
}
```
