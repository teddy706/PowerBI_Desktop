// DIM_Service, DIM_Department, DIM_OS — 3개 쿼리 각각 아래 패턴으로 생성.
// (원본에 없는 파생 차원 테이블이므로 실데이터 전환 후에도 이 3개는 그대로 유지)

// === DIM_Service ===
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "차원테이블_20260828.xlsx"), null, true),
    Sheet = Source{[Item="DIM_Service", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"ServiceKey", type text}, {"표준서비스", type text}, {"표준서비스명", type text},
        {"단위서비스명", type text}, {"서비스등급대표값", type text}
    })
in
    Typed

// === DIM_Department ===
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "차원테이블_20260828.xlsx"), null, true),
    Sheet = Source{[Item="DIM_Department", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"DepartmentKey", type text}, {"부서명", type text}})
in
    Typed

// === DIM_OS ===
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "차원테이블_20260828.xlsx"), null, true),
    Sheet = Source{[Item="DIM_OS", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"OSKey", type text}, {"OS계열", type text}, {"OS명", type text},
        {"버전", type text}, {"EOS대상여부", type text}
    })
in
    Typed
