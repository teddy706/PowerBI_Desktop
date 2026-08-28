let
    LoadSheet = (fileName as text) as table =>
        let
            Source = Excel.Workbook(File.Contents(SourceFolder & fileName), null, true),
            Sheet = Source{[Item="장비목록", Kind="Sheet"]}[Data],
            Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true])
        in
            Promoted,
    Wired = LoadSheet("장비 목록_유선_20260828.xlsx"),
    Wireless = LoadSheet("장비 목록_무선_20260828.xlsx"),
    Combined = Table.Combine({Wired, Wireless}),
    Typed = Table.TransformColumnTypes(Combined, {
        {"AssetKey", type text}, {"SourceType", type text}, {"장비그룹", type text},
        {"장비명", type text}, {"장비IP", type text}, {"장비분류", type text}, {"장비유형", type text},
        {"서비스구분", type text}, {"자산관리자", type text}, {"OS계열", type text}, {"OS", type text}, {"버전", type text},
        {"보안기능", type text}, {"동기화사용", type text}, {"동기화그룹명", type text}, {"장비설명", type text},
        {"미접속일수", Int64.Type}, {"최근접속일자", type date}, {"등록방법", type text}, {"등록일", type datetime}
    })
in
    Typed
