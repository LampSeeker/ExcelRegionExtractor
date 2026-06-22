# Agent Guide

## Project Summary

이 프로젝트는 Excel workbook의 각 sheet에서 **정보영역 bounding box**를 range 문자열로 추출하는 Python 프로젝트입니다.

주요 목표는 값, 병합 셀, border 구조, 삽입 이미지를 이용해 정보가 있는 영역을 찾고, 결과를 JSON과 선택적인 overlay PNG로 저장하는 것입니다.

이 프로젝트는 다음을 하지 않습니다.

- semantic relation 추론
- role hint 생성
- layout root 생성
- structural tag 생성
- title/body 판단
- 특정 시트명, 좌표, 문구에 의존한 규칙 추가

## Runtime

- Python `>=3.10`
- 주요 의존성:
  - `openpyxl>=3.1.0`
  - `Pillow>=10.0.0`

설치:

```powershell
pip install -e .
```

## Main Commands

전체 시트 실행:

```powershell
python scripts/run_all.py --workbook examples/sample.xlsx --out outputs/all_sheets
```

특정 시트 실행:

```powershell
python scripts/extract_info_regions.py --workbook examples/sample.xlsx --sheet "각형맨홀(특2호)" --out outputs/manhole
```

overlay PNG 없이 JSON 중심으로 실행:

```powershell
python scripts/run_all.py --workbook examples/sample.xlsx --out outputs/all_sheets --no-images
```

테스트:

```powershell
pytest
```

## Important Paths

- `src/excel_info_region/extractor.py`: workbook/sheet 실행 흐름과 최종 region 조합
- `src/excel_info_region/cells.py`: 값/병합 셀 occupied 처리
- `src/excel_info_region/components.py`: connected component, bbox 중복 제거, 기하 helper
- `src/excel_info_region/borders.py`: border shell 보정과 border-contact merge
- `src/excel_info_region/image_regions.py`: sheet image anchor를 region box로 변환
- `src/excel_info_region/runner.py`: workbook 실행, 결과 저장, image 추출, overlay PNG 생성
- `src/excel_info_region/schema.py`: `Box`, range helper
- `src/excel_info_region/config.py`: JSON 설정 로더
- `src/excel_info_region/image_export.py`: embedded image 파일 추출
- `src/excel_info_region/visualize.py`: debug overlay PNG 렌더링
- `src/excel_info_region/raw_drawing.py`: drawing/image anchor 관련 저수준 처리
- `scripts/run_all.py`: 전체 또는 특정 시트 실행 CLI
- `scripts/extract_info_regions.py`: `run_all.py` 호환 wrapper
- `config/default.json`: 추출, border 보정, 시각화, image 추출 설정
- `examples/`: 샘플 workbook
- `outputs/`: 실행 결과물
- `tests/test_smoke.py`: 샘플 workbook smoke test

## Processing Model

처리 흐름은 다음과 같습니다.

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

이미지는 cell occupied와 섞어 병합하지 않습니다. 도면 이미지와 바로 아래 표가 붙어 있어도 하나의 영역으로 과병합하지 않는 것이 중요한 설계 의도입니다.

## Output Schema

현재 최종 sheet JSON은 `runner.py` 기준으로 다음 구조를 사용합니다.

```json
{
  "sheet_name": "U형측구",
  "regions": [
    "A1:P1",
    "A3:F8"
  ],
  "images": [
    {
      "name": "Picture 1",
      "range_ref": "A3:F8",
      "path": "images/IMG001_A3_F8_Picture_1.png"
    }
  ]
}
```

주의:

- 내부 extractor 결과에는 `info_regions`가 사용될 수 있습니다.
- 사용자-facing JSON은 `runner.py`에서 `regions`로 변환됩니다.
- `images`는 이미지 이름, anchor range, 추출 파일 경로만 기록합니다.

대표 출력 구조:

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

## Configuration Notes

기본 설정은 `config/default.json`에 있습니다.

중요 옵션:

- `include_values`: 값이 있는 셀을 occupied로 사용
- `include_merged_cells`: 병합 셀 범위 전체를 occupied로 사용
- `include_images`: image anchor를 region 후보로 고려
- `include_grouped_drawing_images`: grouped drawing image 처리
- `connectivity`: value component 연결성
- `use_borders`: border-aware 보정 사용
- `strong_borders_only`: 강한 border만 사용할지 여부
- `add_border_only_regions`: 값이 없는 border-only 영역을 최종 영역으로 추가할지 여부
- `use_border_contact_merge`: border contact 기반 최종 병합 사용
- `border_contact_min_touched_sides`: 같은 border component에 접촉해야 하는 최소 변 수
- `extract_embedded_images`: embedded image 파일 추출 여부
- `embedded_image_dir`: sheet output 하위 image 저장 폴더명
- `visualization.enabled`: overlay PNG 생성 여부
- `visualization.font_path`: 한글 PNG 렌더링용 폰트 경로

`add_border_only_regions`의 기본값은 `false`입니다. 값이 전혀 없는 테두리 박스만으로 정보영역을 만들지 않는 보수적인 동작입니다.

## Development Rules

- 최종 정보영역은 range 문자열 목록으로 유지합니다.
- 새 규칙은 특정 시트명, 특정 좌표, 특정 텍스트에 의존하지 않아야 합니다.
- 이미지 영역은 cell region과 임의 병합하지 않습니다.
- border 보정은 큰 shell이 제목, 도면, 다른 표까지 삼키지 않도록 면적 비율과 추가 row/col 제한을 유지해야 합니다.
- `openpyxl`은 수식을 계산하지 않습니다. overlay PNG는 `data_only=True` workbook을 사용하므로, Excel에 cached formula result가 없으면 수식 결과가 비어 보일 수 있습니다.
- 한글이 overlay PNG에서 깨지면 `visualization.font_path`에 `C:/Windows/Fonts/malgun.ttf` 같은 폰트를 지정합니다.
- 결과 schema를 바꿀 때는 README, 이 문서, smoke test를 함께 갱신합니다.

## Verification

변경 후 최소 확인:

```powershell
pytest
python scripts/run_all.py --workbook examples/sample.xlsx --out outputs/all_sheets --no-images
```

시각화나 image 추출을 건드렸다면 `--no-images` 없이 실행해서 sheet별 `info_regions.png`와 `images/` 결과도 확인합니다.
