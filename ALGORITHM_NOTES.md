# Algorithm Notes

## 1. Connected Component Labeling

- 입력 신호: 값이 있는 셀, 값이 있는 merged cell의 전체 범위, 이미지 anchor 범위, 선택적으로 강한 테두리 셀.
- 역할: 가장 작은 primitive block 생성.
- 한계: 빈 행/빈 열이 있으면 같은 의미영역도 끊어진다.

## 2. Projection Profile

- 입력 신호: 행별/열별 occupied, non-empty, numeric, formula, bold, border count.
- 역할: split 후보가 되는 blank row/column gap을 찾는다.
- 한계: 빈 gap만으로는 같은 정보영역인지 다른 정보영역인지 확정할 수 없다.

## 3. Region Growing / Region Merging

- 입력: connected component primitive blocks.
- 방식: 가장 점수가 높은 인접 블록을 반복 병합한다.
- 장점: `B16:J45`와 `B49:J53`처럼 중간에 빈 행이 있어도 같은 폭/스타일/테두리 신호가 있으면 묶을 수 있다.

## 4. Graph-based Clustering + Union-Find

- 입력: primitive block graph.
- 방식: block 간 pairwise edge score를 계산하고 threshold 이상이면 union한다.
- 장점: 하드코딩된 텍스트 조건 없이 다중 신호 점수로 grouping한다.

## 5. Hierarchical Agglomerative Clustering

- 입력: primitive blocks.
- 방식: 가장 비슷한 두 region을 병합하고 feature를 다시 계산한다.
- 장점: 병합 순서와 score를 추적하기 좋다.
- 주의: threshold가 낮으면 과병합될 수 있다.

## 6. DBSCAN-like Spatial Clustering

- 입력: primitive block의 중심 좌표, density, numeric density 등.
- 방식: 거리가 가까운 block을 같은 cluster로 묶는다.
- 장점: 빠르게 큰 군집을 만들 수 있다.
- 한계: 엑셀에서는 거리만으로 의미영역을 판단하기 어려우므로 보조 실험용이다.

## 7. XY-cut Recursive Segmentation

- 입력: 전체 effective bounds와 occupied grid.
- 방식: 가장 큰 빈 행/열 gap을 찾아 재귀적으로 영역을 나눈다.
- 장점: Root 내부 Candidate split에 유용하다.
- 한계: 전체 root grouping보다는 세부 분할 후보 탐색에 더 적합하다.

## 추천 사용 방식

1. Connected Component로 primitive block 생성.
2. Projection Profile로 blank gap과 section anchor 확인.
3. Graph Union-Find 또는 Region Growing으로 RootRegion grouping.
4. XY-cut으로 RootRegion 내부 CandidateRegion split.
5. VLM은 후보 영역의 의미 판단/좌표 검증에만 사용.
