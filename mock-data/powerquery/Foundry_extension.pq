// Palantir Foundry 연계 확장: 이상징후 판정 결과 + 화이트리스트 룰 마스터 현황.
// Foundry의 실시간 명령어 교차분석·1차/2차 판정 입력 UI 자체는 Power BI 영역이
// 아니므로 구현하지 않고, 그 "결과"를 관리 대시보드로 보여주는 부분만 반영한다.

// === FACT_JudgmentCase (Foundry_연계_20260828.xlsx, "판정이력" 시트) ===
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "Foundry_연계_20260828.xlsx"), null, true),
    Sheet = Source{[Item="판정이력", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"CaseID", type text}, {"케이스유형", type text}, {"HostID", type text}, {"AssetKey", type text},
        {"탐지일시", type date}, {"1차판단자", type text}, {"1차소명여부", type text}, {"1차소명일", type date},
        {"2차판단자", type text}, {"2차판정", type text}, {"2차판정일", type date}, {"처리소요일", Int64.Type}
    })
in
    Typed

// === DIM_WhitelistRule (Foundry_연계_20260828.xlsx, "화이트리스트룰" 시트) ===
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "Foundry_연계_20260828.xlsx"), null, true),
    Sheet = Source{[Item="화이트리스트룰", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"RuleID", type text}, {"명령어패턴", type text}, {"등록근거_CaseID", type text},
        {"등록일", type date}, {"반복발생횟수", Int64.Type}, {"재검토대상여부", type text}, {"재검토사유", type text}
    })
in
    Typed
