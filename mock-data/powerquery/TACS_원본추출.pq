// TACS 시스템 자체의 원본 추출 데이터. 서버관리(FACT_Host)와는 독립된 시스템이므로
// 이름/IP 표기가 다를 수 있고, 서버관리가 "등록 Y"라 주장해도 이 시스템에 행이
// 아예 없을 수 있다 (완전 미등록 케이스).
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "지역NW운용본부_20260828.xlsx"), null, true),
    Sheet = Source{[Item="TACS원본추출", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"HostID", type text}, {"장비명", type text}, {"IP", type text}, {"등록일시", type datetime}
    })
in
    Typed
