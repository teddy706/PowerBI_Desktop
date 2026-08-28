# Power Query 정제 계층 (2단계 결과물)

이 폴더의 `.pq` 파일들은 Power BI Desktop의 Power Query 편집기(고급 편집기)에 그대로 붙여넣을 수 있는 M 코드입니다. 3단계(시맨틱 모델 구축)에서 Modeling MCP로 실제 모델에 반영합니다.

## 로드 순서
1. `00_SourceFolder_Parameter.pq` → 매개변수 `SourceFolder` 먼저 생성
2. `FACT_Host.pq`, `FACT_Asset.pq`, `FACT_TACS.pq`, `FACT_SecurityCompliance.pq`, `DIM_tables.pq` (순서 무관, 서로 독립)
3. `FACT_TACS_Matching.pq` → **FACT_Asset, FACT_Host가 먼저 로드되어 있어야 함** (내부에서 두 쿼리를 참조)

## 검증된 매칭 분포
`FACT_TACS_Matching`의 로직을 파이썬으로 동일하게 재현해 mock 데이터 기준으로 사전 검증함 (`../generate_mock_data.py` 실행 시 `../summary.txt`에 기록):
MATCH 80% / NAME_MATCH 5% / IP_MATCH 5% / NO_MATCH 10% — 계획서 3-5절 목표치와 일치.

## 실데이터 전환 시 체크리스트
- `SourceFolder` 매개변수 값만 실제 원본 폴더 경로로 교체
- 원본 파일명이 mock과 다르면 각 쿼리의 파일명 리터럴(`"호스트명조회_20260828.xlsx"` 등)을 실제 파일명으로 수정
- `FACT_TACS.pq`는 "보안통제"를 제외한 모든 시트를 자동으로 합치도록 만들어서, 실원본에 지역 시트가 여러 개(서부/동부/남부 등)여도 쿼리 수정 없이 동작함
- 키 컬럼(HostID, 장비명, 장비IP, 대표IP) 스펠링이 실원본과 정확히 일치하는지 최종 확인 (계획서 8장 핵심 유의사항)
