// 원본에 존재하지 않는 계산 테이블. FACT_Host(서버관리, "TACS 등록됨"이라는 주장)를
// TACS_원본추출(TACS 시스템 실제 데이터)과 HostID로 조인해, 주장이 실제로
// 뒷받침되는지 검증한다. FACT_Host, TACS_원본추출 쿼리가 먼저 로드되어 있어야 함.
let
    Merged = Table.NestedJoin(FACT_Host, {"HostID"}, TACS_원본추출, {"HostID"}, "TACS", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, "TACS", {"HostID", "장비명", "IP"}, {"TACS_HostID", "TACS원본_장비명", "TACS원본_IP"}),

    Classified = Table.AddColumn(Expanded, "VerifyStatus", each
        if [TACS_HostID] <> null then
            if [호스트명] = [TACS원본_장비명] and [대표IP] = [TACS원본_IP] then "VERIFIED"
            else if [호스트명] = [TACS원본_장비명] then "VERIFIED_NAME_ONLY"
            else if [대표IP] = [TACS원본_IP] then "VERIFIED_IP_ONLY"
            else "UNVERIFIED"
        else if [서버관리_TACS등록여부] = "Y" then "UNVERIFIED"
        else null,
        type text),

    WithReason = Table.AddColumn(Classified, "불일치사유", each
        if [VerifyStatus] = "VERIFIED_NAME_ONLY" then "IP_MISMATCH"
        else if [VerifyStatus] = "VERIFIED_IP_ONLY" then "NAME_MISMATCH"
        else if [VerifyStatus] = "UNVERIFIED" then "NOT_FOUND"
        else null,
        type text),

    WithScore = Table.AddColumn(WithReason, "VerifyScore", each
        if [VerifyStatus] = "VERIFIED" then 1.0
        else if [VerifyStatus] = "VERIFIED_NAME_ONLY" or [VerifyStatus] = "VERIFIED_IP_ONLY" then 0.5
        else if [VerifyStatus] = "UNVERIFIED" then 0.0
        else null,
        type number),

    Final = Table.SelectColumns(WithScore,
        {"HostID", "호스트명", "대표IP", "서버관리_TACS등록여부", "TACS원본_장비명", "TACS원본_IP", "VerifyStatus", "불일치사유", "VerifyScore"})
in
    Final
