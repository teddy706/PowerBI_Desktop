// 원본은 "서부" 외에도 지역별 시트(동부/남부 등)가 더 있을 수 있음.
// "보안통제" 시트를 제외한 모든 시트를 자동으로 합쳐서, 실데이터에 지역 시트가
// 몇 개든 쿼리 수정 없이 그대로 동작하도록 설계.
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "지역NW운용본부_20260828.xlsx"), null, true),
    RegionSheets = Table.SelectRows(Source, each [Kind] = "Sheet" and [Name] <> "보안통제"),
    PromotedTables = List.Transform(RegionSheets[Data], each Table.PromoteHeaders(_, [PromoteAllScalars=true])),
    Combined = Table.Combine(PromotedTables),
    Typed = Table.TransformColumnTypes(Combined, {
        {"HostID", type text}, {"호스트명", type text}, {"장치ID", type text}, {"운영부서", type text},
        {"표준서비스명", type text}, {"단위서비스명", type text},
        {"TACS수용(OS)", type text}, {"TACS수용(DB)", type text},
        {"TACS수용(OS)예외사유", type text}, {"TACS수용(DB)예외사유", type text},
        {"운영상태", type text}, {"설비상태", type text}, {"VM여부", type text}
    })
in
    Typed
