let
    Source = Excel.Workbook(File.Contents(SourceFolder & "지역NW운용본부_20260828.xlsx"), null, true),
    Sheet = Source{[Item="보안통제", Kind="Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"HostID", type text}, {"웹쉘Agent설치", type text}, {"웹쉘정책설정", type text},
        {"서버백신설치", type text}, {"서버백신실시간감시", type text}, {"서버백신HIPS", type text}, {"서버백신FullScan", type text},
        {"EDR설치", type text}, {"EDRFullScan", type text},
        {"중앙로그관리수용", type text}, {"로그보관기간설정", type text}, {"ACL설정", type text},
        {"EOS해소여부", type text}, {"SMP진단", type text}, {"SMP조치", type text},
        {"개인정보암호화", type text}, {"개인정보암호화일정", type text}
    })
in
    Typed
