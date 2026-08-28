// 10장 확장 4개 테이블의 M 코드. 각각 Power BI Desktop에서 별도 쿼리로 붙여넣는다.

// === DIM_TACS_Scope_Rule (차원테이블_20260828.xlsx, 4번째 시트) ===
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "차원테이블_20260828.xlsx"), null, true),
    Sheet = Source{[Item="DIM_TACS_Scope_Rule", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"RuleKey", type text}, {"적용조건_서비스등급", type text}, {"적용조건_OS계열", type text},
        {"적용조건_서비스구분", type text}, {"대상여부", type text}, {"규칙설명", type text}, {"최종개정일", type date}
    })
in
    Typed

// === FACT_TACS_Exception (TACS_관리대장_20260828.xlsx, "예외관리" 시트) ===
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "TACS_관리대장_20260828.xlsx"), null, true),
    Sheet = Source{[Item="예외관리", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"ExceptionKey", type text}, {"HostID", type text}, {"예외사유", type text}, {"예외사유상세", type text},
        {"승인자", type text}, {"승인일", type date}, {"재검토주기_개월", Int64.Type}, {"만료예정일", type date}
    })
in
    Typed

// === FACT_TACS_History (TACS_관리대장_20260828.xlsx, "이력관리" 시트) ===
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "TACS_관리대장_20260828.xlsx"), null, true),
    Sheet = Source{[Item="이력관리", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"HistoryKey", type text}, {"HostID", type text}, {"스냅샷일자", type date},
        {"TACS수용_OS", type text}, {"TACS수용_DB", type text}, {"VerifyStatus", type text},
        {"변경유형", type text}, {"이전값", type text}, {"변경일", type date}
    })
in
    Typed

// === FACT_TACS_Escalation (TACS_관리대장_20260828.xlsx, "에스컬레이션" 시트) ===
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "TACS_관리대장_20260828.xlsx"), null, true),
    Sheet = Source{[Item="에스컬레이션", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"EscalationKey", type text}, {"HostID", type text}, {"리스크등급", type text},
        {"발생일", type date}, {"SLA기한", type date}, {"담당자", type text},
        {"처리상태", type text}, {"최종조치일", type date}
    })
in
    Typed
