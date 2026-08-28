// 원본에 존재하지 않는 계산 테이블. FACT_Asset(장비목록)과 FACT_Host(호스트조회)를
// 1차: 장비명=호스트명, 2차: 장비IP=대표IP 기준으로 비교해 정합성을 검증한다.
// FACT_Asset, FACT_Host 쿼리가 이미 로드되어 있어야 함 (둘 다 이 파일보다 먼저 로드).
let
    // 1차 매칭: 장비명 = 호스트명
    MergedByName = Table.NestedJoin(FACT_Asset, {"장비명"}, FACT_Host, {"호스트명"}, "HostByName", JoinKind.LeftOuter),
    ExpandedByName = Table.ExpandTableColumn(MergedByName, "HostByName",
        {"HostID", "호스트명", "대표IP"}, {"NameMatch_HostID", "NameMatch_호스트명", "NameMatch_대표IP"}),

    // 2차 매칭: 장비IP = 대표IP
    MergedByIP = Table.NestedJoin(ExpandedByName, {"장비IP"}, FACT_Host, {"대표IP"}, "HostByIP", JoinKind.LeftOuter),
    ExpandedByIP = Table.ExpandTableColumn(MergedByIP, "HostByIP",
        {"HostID", "호스트명", "대표IP"}, {"IPMatch_HostID", "IPMatch_호스트명", "IPMatch_대표IP"}),

    // 판정: 이름/IP 매칭이 같은 HostID를 가리키면 MATCH, 이름만 맞으면 NAME_MATCH,
    //       IP만 맞으면 IP_MATCH, 둘 다 없으면 NO_MATCH
    Classified = Table.AddColumn(ExpandedByIP, "MatchStatus", each
        if [NameMatch_HostID] <> null and [IPMatch_HostID] <> null and [NameMatch_HostID] = [IPMatch_HostID] then "MATCH"
        else if [NameMatch_HostID] <> null then "NAME_MATCH"
        else if [IPMatch_HostID] <> null then "IP_MATCH"
        else "NO_MATCH",
        type text),

    WithMatchScore = Table.AddColumn(Classified, "MatchScore", each
        if [MatchStatus] = "MATCH" then 1.0
        else if [MatchStatus] = "NAME_MATCH" or [MatchStatus] = "IP_MATCH" then 0.5
        else 0.0,
        type number),

    WithHostID = Table.AddColumn(WithMatchScore, "HostID", each
        if [NameMatch_HostID] <> null then [NameMatch_HostID] else [IPMatch_HostID],
        type text),
    WithHostName = Table.AddColumn(WithHostID, "호스트명", each
        if [NameMatch_호스트명] <> null then [NameMatch_호스트명] else [IPMatch_호스트명],
        type text),
    WithRepIP = Table.AddColumn(WithHostName, "대표IP", each
        if [IPMatch_대표IP] <> null then [IPMatch_대표IP] else [NameMatch_대표IP],
        type text),

    Final = Table.SelectColumns(WithRepIP,
        {"AssetKey", "장비명", "장비IP", "HostID", "호스트명", "대표IP", "MatchStatus", "MatchScore"})
in
    Final
