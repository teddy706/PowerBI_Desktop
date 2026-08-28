let
    Source = Excel.Workbook(File.Contents(SourceFolder & "호스트조회_20260828.xlsx"), null, true),
    Sheet = Source{[Item="호스트조회", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"HostID", type text}, {"호스트명", type text}, {"대표IP", type text}, {"대표IP노마스킹", type text},
        {"OS", type text}, {"OS계열", type text}, {"OS버전", type text}, {"Platform", type text},
        {"CPU스펙", type text}, {"CPU유형", type text},
        {"물리적CPU수", Int64.Type}, {"CPU물리Core수", Int64.Type}, {"CPU논리Core수", Int64.Type},
        {"메모리용량MB", Int64.Type},
        {"호스트운용부서", type text}, {"호스트운용자", type text},
        {"PO부서", type text}, {"PO", type text}, {"BA부서", type text}, {"BA", type text},
        {"표준서비스", type text}, {"표준서비스명", type text}, {"단위서비스", type text}, {"단위서비스명", type text},
        {"서비스등급", type text}, {"VM수집방식상세", type text}, {"설비상태", type text},
        {"최근접속일자", type date}, {"미접속일수", Int64.Type}, {"서버관리_TACS등록여부", type text}
    })
in
    Typed
