"""
Mock data generator for TACS MVP project.
Reproduces the ORIGINAL 4-file structure (2 domains split across files, 3rd domain in one multi-sheet file)
so that Power Query logic built against this mock data transfers to real source files unchanged.

Run: python generate_mock_data.py
Output: ./out/*.xlsx  (원본과 동일한 파일명 패턴, 시트 구조)
"""
import random
from datetime import date, timedelta
from pathlib import Path
from openpyxl import Workbook

random.seed(42)  # 재현 가능한 목데이터 (재실행해도 같은 결과)

OUT_DIR = Path(__file__).parent / "out"
OUT_DIR.mkdir(exist_ok=True)
TODAY = date(2026, 8, 28)

DEPARTMENTS = ["NW운용1팀", "NW운용2팀", "IDC운용팀", "Cloud사업팀"]
PO_DEPTS = ["Cloud사업팀", "Platform사업팀", "Infra사업팀"]
BA_DEPTS = ["BA1", "BA2", "BA3"]

STD_SERVICES = [
    ("SVC01", "KT Cloud", "KTCLD"),
    ("SVC02", "Cloud DB", "KTDB"),
    ("SVC03", "IPTV", "KTIPTV"),
    ("SVC04", "GiGAeyes", "KTGE"),
    ("SVC05", "AI Platform", "KTAI"),
    ("SVC06", "BCP", "KTBCP"),
]
UNIT_SERVICES = {
    "KT Cloud": ["Cloud Compute", "Cloud Storage"],
    "Cloud DB": ["Cloud DB-Prod", "Cloud DB-Dev"],
    "IPTV": ["IPTV-Head", "IPTV-VOD"],
    "GiGAeyes": ["GiGAeyes-Core", "GiGAeyes-Edge"],
    "AI Platform": ["AI-Training", "AI-Serving"],
    "BCP": ["BCP-DR", "BCP-Backup"],
}
GRADE_POOL = ["Critical"] * 15 + ["High"] * 30 + ["Medium"] * 40 + ["Low"] * 15

LINUX_OS = [("Linux", "Rocky Linux", "9.4"), ("Linux", "Rocky Linux", "8.9"),
            ("Linux", "RHEL", "8.8"), ("Linux", "RHEL", "7.9"),
            ("Linux", "Ubuntu", "22.04"), ("Linux", "Ubuntu", "20.04")]
WINDOWS_OS = [("Windows", "Windows Server", "2022"), ("Windows", "Windows Server", "2019"),
              ("Windows", "Windows Server", "2012 R2")]
EOS_VERSIONS = {("RHEL", "7.9"), ("Windows Server", "2012 R2")}

EXCEPTION_REASONS = ["Legacy OS"] * 40 + ["Vendor 제한"] * 25 + ["서비스 영향도"] * 20 + ["업그레이드 예정"] * 15


def rand_date(start: date, end: date) -> date:
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, max(delta, 0)))


def pick_weighted(pool):
    return random.choice(pool)


# ---------------------------------------------------------------------------
# FACT_Host  (500건) — 모든 것의 기준키
# ---------------------------------------------------------------------------
hosts = []
name_seq = {prefix: 0 for _, _, prefix in STD_SERVICES}
ip_counter = 0

for i in range(1, 501):
    host_id = f"HOST{i:05d}"
    std_code, std_name, prefix = random.choice(STD_SERVICES)
    name_seq[prefix] += 1
    host_name = f"{prefix}{name_seq[prefix]:03d}"

    unit_name = random.choice(UNIT_SERVICES[std_name])
    grade = pick_weighted(GRADE_POOL)

    is_linux = random.random() < 0.70
    os_family, os_name, os_ver = random.choice(LINUX_OS if is_linux else WINDOWS_OS)

    is_vm = random.random() < 0.80
    platform = "VMware" if is_vm else "Physical"
    cpu_type = "Intel" if random.random() < 0.80 else "AMD"
    phys_cpu = random.choice([1, 2])
    phys_core = random.choice([8, 16, 24, 32])
    logi_core = phys_core * 2
    mem_mb = random.choice([16384, 32768, 65536, 131072])

    ip_counter += 1
    rep_ip = f"10.{20 + STD_SERVICES.index((std_code, std_name, prefix))}.{(ip_counter // 254) + 1}.{(ip_counter % 254) + 1}"

    op_dept = random.choice(DEPARTMENTS)
    po_dept = random.choice(PO_DEPTS)
    ba_dept = random.choice(BA_DEPTS)

    stale_days = random.choice([0, 0, 0, 1, 2, 3, 15, 30, 60]) if random.random() < 0.95 else random.randint(30, 90)
    last_seen = TODAY - timedelta(days=stale_days)

    facility_status = random.choices(["운영중", "점검중", "폐기예정"], weights=[90, 5, 5])[0]
    vm_collect = ("Agent" if random.random() < 0.9 else "Agentless") if is_vm else "-"

    hosts.append({
        "HostID": host_id,
        "호스트명": host_name,
        "대표IP": rep_ip,
        "대표IP노마스킹": rep_ip,
        "OS": os_name,
        "OS계열": os_family,
        "OS버전": os_ver,
        "Platform": platform,
        "CPU스펙": f"{cpu_type} Xeon/EPYC Series",
        "CPU유형": cpu_type,
        "물리적CPU수": phys_cpu,
        "CPU물리Core수": phys_core,
        "CPU논리Core수": logi_core,
        "메모리용량MB": mem_mb,
        "호스트운용부서": op_dept,
        "호스트운용자": f"운영자{random.randint(1, 30):02d}",
        "PO부서": po_dept,
        "PO": f"PO{random.randint(1, 15):02d}",
        "BA부서": ba_dept,
        "BA": f"BA사원{random.randint(1, 15):02d}",
        "표준서비스": std_code,
        "표준서비스명": std_name,
        "단위서비스": f"{std_code}-{UNIT_SERVICES[std_name].index(unit_name)+1}",
        "단위서비스명": unit_name,
        "서비스등급": grade,
        "VM수집방식상세": vm_collect,
        "설비상태": facility_status,
        "최근접속일자": last_seen,
        "미접속일수": stale_days,
    })

# ---------------------------------------------------------------------------
# DIM_Service / DIM_Department / DIM_OS — Host에서 distinct 추출
# ---------------------------------------------------------------------------
dim_service_rows = []
skey = 0
for std_code, std_name, _ in STD_SERVICES:
    for unit_name in UNIT_SERVICES[std_name]:
        skey += 1
        grades_used = sorted({h["서비스등급"] for h in hosts if h["표준서비스명"] == std_name and h["단위서비스명"] == unit_name}) or ["Medium"]
        dim_service_rows.append({
            "ServiceKey": f"SVCKEY{skey:03d}",
            "표준서비스": std_code, "표준서비스명": std_name,
            "단위서비스명": unit_name, "서비스등급대표값": grades_used[0],
        })

dim_dept_rows = []
dkey = 0
for d in sorted(set(DEPARTMENTS) | set(PO_DEPTS) | set(BA_DEPTS)):
    dkey += 1
    dim_dept_rows.append({"DepartmentKey": f"DEPTKEY{dkey:03d}", "부서명": d})

dim_os_rows = []
okey = 0
for fam, name, ver in sorted(set((h["OS계열"], h["OS"], h["OS버전"]) for h in hosts)):
    okey += 1
    dim_os_rows.append({
        "OSKey": f"OSKEY{okey:03d}", "OS계열": fam, "OS명": name, "버전": ver,
        "EOS대상여부": "Y" if (name, ver) in EOS_VERSIONS else "N",
    })

# ---------------------------------------------------------------------------
# FACT_Asset (100건) — 유선/무선 분리, 원본 파일 예시(KT-L4-SEOUL-01 등) 준수
# ---------------------------------------------------------------------------
REGIONS = ["SEOUL", "BUSAN", "DAEJEON", "GWANGJU", "INCHEON"]
WIRED_TYPES = ["Router", "Switch", "Firewall", "L4", "VPN"]
WIRELESS_TYPES = ["AP", "WLC", "Wireless-Bridge"]
TYPE_ABBR = {"Router": "RT", "Switch": "SW", "Firewall": "FW", "L4": "L4", "VPN": "VPN",
             "AP": "AP", "WLC": "WLC", "Wireless-Bridge": "WBR"}


def perturb_ip(ip: str) -> str:
    parts = ip.split(".")
    old_last = int(parts[-1])
    new_last = old_last
    while new_last == old_last:
        new_last = random.randint(1, 254)
    parts[-1] = str(new_last)
    return ".".join(parts)


def perturb_name(name: str) -> str:
    import re
    m = re.match(r"^([A-Za-z]+)(\d+)$", name)
    if not m:
        return name + "-X"
    letters, digits = m.groups()
    cut = max(len(letters) - 2, 1)
    return f"{letters[:cut]}-{letters[cut:]}{digits}"


all_asset_rows = []
for i in range(1, 101):
    wired = i <= 70  # 70 유선 / 30 무선
    dev_type = random.choice(WIRED_TYPES if wired else WIRELESS_TYPES)
    region = random.choice(REGIONS)
    seq = (i if wired else i - 70)
    device_name = f"KT-{TYPE_ABBR[dev_type]}-{region}-{seq:02d}"
    device_ip = f"10.10.{random.randint(1, 20)}.{random.randint(1, 254)}"
    is_linux = random.random() < 0.6
    os_family = "Linux" if is_linux else "Network OS"
    os_name, os_ver = ("Rocky Linux", "9.4") if is_linux else (random.choice(["IOS-XE", "JunOS", "PAN-OS"]), f"{random.randint(12,17)}.{random.randint(0,9)}")
    stale = random.choice([0, 1, 2, 5, 10]) if random.random() < 0.9 else random.randint(15, 45)

    row = {
        "AssetKey": f"ASSET{i:04d}",
        "SourceType": "유선" if wired else "무선",
        "장비그룹": random.choice(["네트워크부문", "보안부문", "인프라부문"]),
        "장비명": device_name,
        "장비IP": device_ip,
        "장비분류": "NETWORK",
        "장비유형": dev_type,
        "서비스구분": "운영" if random.random() < 0.85 else "개발",
        "자산관리자": random.choice(DEPARTMENTS),
        "OS계열": os_family,
        "OS": os_name,
        "버전": os_ver,
        "보안기능": "적용" if random.random() < 0.8 else "미적용",
        "동기화사용": "Y" if random.random() < 0.85 else "N",
        "동기화그룹명": f"NTP-GROUP-{random.randint(1,3):02d}",
        "장비설명": f"{region} IDC {dev_type}",
        "미접속일수": stale,
        "최근접속일자": TODAY - timedelta(days=stale),
        "등록방법": "자동" if random.random() < 0.9 else "수동",
        "등록일": rand_date(date(2024, 1, 1), date(2026, 6, 30)),
        "_wired": wired,
    }
    all_asset_rows.append(row)

# --- FACT_Host와의 겹침 주입: FACT_TACS_Matching이 100건 기준 80/5/5/10을 갖도록 ---
# (겹침 없이는 Power Query 매칭 로직이 항상 NO_MATCH만 내놓아 검증 대시보드가 무의미해짐)
overlap_hosts = random.sample(hosts, 90)
match_hosts, name_match_hosts, ip_match_hosts = overlap_hosts[:80], overlap_hosts[80:85], overlap_hosts[85:90]
random.shuffle(all_asset_rows)  # 어느 인덱스가 겹칠지 유선/무선에 고르게 섞이도록

for row, h in zip(all_asset_rows[0:80], match_hosts):
    row["장비명"], row["장비IP"] = h["호스트명"], h["대표IP"]
for row, h in zip(all_asset_rows[80:85], name_match_hosts):
    row["장비명"], row["장비IP"] = h["호스트명"], perturb_ip(h["대표IP"])
for row, h in zip(all_asset_rows[85:90], ip_match_hosts):
    row["장비명"], row["장비IP"] = perturb_name(h["호스트명"]), h["대표IP"]
# 나머지 all_asset_rows[90:100] 은 원래 생성된 독립 네트워크 장비명 그대로 → NO_MATCH

assets_wired = [{k: v for k, v in r.items() if k != "_wired"} for r in all_asset_rows if r["_wired"]]
assets_wireless = [{k: v for k, v in r.items() if k != "_wired"} for r in all_asset_rows if not r["_wired"]]

# ---------------------------------------------------------------------------
# FACT_TACS (500건, Host와 1:1) + FACT_SecurityCompliance (500건, Host와 1:1)
# 원본은 '지역NW운용본부' 한 파일 안에 시트로 분리되어 있음
# ---------------------------------------------------------------------------
tacs_rows, sec_rows = [], []
for idx, h in enumerate(hosts, start=1):
    os_accept = "Y" if random.random() < 0.90 else "N"
    db_accept = "Y" if random.random() < 0.85 else "N"
    os_reason = "" if os_accept == "Y" else pick_weighted(EXCEPTION_REASONS)
    db_reason = "" if db_accept == "Y" else pick_weighted(EXCEPTION_REASONS)

    tacs_rows.append({
        "HostID": h["HostID"], "호스트명": h["호스트명"], "장치ID": f"DEV{idx:05d}",
        "운영부서": h["호스트운용부서"], "표준서비스명": h["표준서비스명"], "단위서비스명": h["단위서비스명"],
        "TACS수용(OS)": os_accept, "TACS수용(DB)": db_accept,
        "TACS수용(OS)예외사유": os_reason, "TACS수용(DB)예외사유": db_reason,
        "운영상태": h["설비상태"], "설비상태": h["설비상태"],
        "VM여부": "Y" if h["Platform"] == "VMware" else "N",
    })

    is_eos = (h["OS"], h["OS버전"]) in EOS_VERSIONS
    webshell_agent = random.random() < 0.85
    av_installed = random.random() < 0.92
    edr_installed = random.random() < 0.78
    central_log = random.random() < 0.88
    smp_diag = random.random() < 0.80
    privacy_enc = random.random() < 0.75

    def yn(b):
        return "Y" if b else "N"

    sec_rows.append({
        "HostID": h["HostID"],
        "웹쉘Agent설치": yn(webshell_agent),
        "웹쉘정책설정": yn(webshell_agent and random.random() < 0.85),
        "서버백신설치": yn(av_installed),
        "서버백신실시간감시": yn(av_installed and random.random() < 0.90),
        "서버백신HIPS": yn(av_installed and random.random() < 0.70),
        "서버백신FullScan": yn(av_installed and random.random() < 0.80),
        "EDR설치": yn(edr_installed),
        "EDRFullScan": yn(edr_installed and random.random() < 0.75),
        "중앙로그관리수용": yn(central_log),
        "로그보관기간설정": random.choice(["90일", "180일", "365일"]) if central_log else "미설정",
        "ACL설정": yn(random.random() < 0.82),
        "EOS해소여부": yn(random.random() < (0.40 if is_eos else 0.95)),
        "SMP진단": yn(smp_diag),
        "SMP조치": yn(smp_diag and random.random() < 0.70),
        "개인정보암호화": yn(privacy_enc),
        "개인정보암호화일정": str(rand_date(date(2025, 1, 1), TODAY)) if privacy_enc else str(rand_date(TODAY, date(2027, 1, 1))),
    })


def write_sheet(ws, rows):
    if not rows:
        return
    headers = list(rows[0].keys())
    ws.append(headers)
    for r in rows:
        ws.append([r[h] for h in headers])


def save_single(filename, sheet_name, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    write_sheet(ws, rows)
    wb.save(OUT_DIR / filename)


# 1) 장비 목록_유선 / 무선
save_single("장비 목록_유선_20260828.xlsx", "장비목록", assets_wired)
save_single("장비 목록_무선_20260828.xlsx", "장비목록", assets_wireless)

# 2) 호스트조회
save_single("호스트조회_20260828.xlsx", "호스트조회", hosts)

# 3) 지역NW운용본부 (멀티시트: TACS현황 + 보안통제)
wb = Workbook()
ws1 = wb.active
ws1.title = "서부"
write_sheet(ws1, tacs_rows)
ws2 = wb.create_sheet("보안통제")
write_sheet(ws2, sec_rows)
wb.save(OUT_DIR / "지역NW운용본부_20260828.xlsx")

# 4) 차원 테이블 (Power BI에서 직접 로드할 참조 데이터 — 원본에 없는 파생 테이블)
wb = Workbook()
ws = wb.active
ws.title = "DIM_Service"
write_sheet(ws, dim_service_rows)
ws2 = wb.create_sheet("DIM_Department")
write_sheet(ws2, dim_dept_rows)
ws3 = wb.create_sheet("DIM_OS")
write_sheet(ws3, dim_os_rows)
wb.save(OUT_DIR / "차원테이블_20260828.xlsx")

# ---------------------------------------------------------------------------
# 매칭 검증 시뮬레이션 (Power Query 매칭 로직을 파이썬으로 미리 재현해 분포 확인)
# ---------------------------------------------------------------------------
host_by_name = {h["호스트명"]: h for h in hosts}
host_by_ip = {h["대표IP"]: h for h in hosts}
match_counts = {"MATCH": 0, "NAME_MATCH": 0, "IP_MATCH": 0, "NO_MATCH": 0}
for a in all_asset_rows:
    h_by_name = host_by_name.get(a["장비명"])
    h_by_ip = host_by_ip.get(a["장비IP"])
    if h_by_name and h_by_ip and h_by_name["HostID"] == h_by_ip["HostID"]:
        match_counts["MATCH"] += 1
    elif h_by_name:
        match_counts["NAME_MATCH"] += 1
    elif h_by_ip:
        match_counts["IP_MATCH"] += 1
    else:
        match_counts["NO_MATCH"] += 1

# ---------------------------------------------------------------------------
# 요약 리포트 (콘솔 한글 인코딩 문제 회피를 위해 UTF-8 파일로 기록)
# ---------------------------------------------------------------------------
with open(Path(__file__).parent / "summary.txt", "w", encoding="utf-8") as f:
    f.write(f"FACT_Host: {len(hosts)}건\n")
    f.write(f"  - Linux/Windows: {sum(1 for h in hosts if h['OS계열']=='Linux')}/{sum(1 for h in hosts if h['OS계열']=='Windows')}\n")
    f.write(f"  - VM/Physical: {sum(1 for h in hosts if h['Platform']=='VMware')}/{sum(1 for h in hosts if h['Platform']=='Physical')}\n")
    f.write(f"FACT_Asset: 유선 {len(assets_wired)}건, 무선 {len(assets_wireless)}건 (합계 {len(all_asset_rows)}건)\n")
    f.write(f"FACT_TACS: {len(tacs_rows)}건, OS수용Y {sum(1 for t in tacs_rows if t['TACS수용(OS)']=='Y')}건 ({sum(1 for t in tacs_rows if t['TACS수용(OS)']=='Y')/len(tacs_rows):.0%})\n")
    f.write(f"FACT_SecurityCompliance: {len(sec_rows)}건, EDR설치Y {sum(1 for s in sec_rows if s['EDR설치']=='Y')}건 ({sum(1 for s in sec_rows if s['EDR설치']=='Y')/len(sec_rows):.0%})\n")
    f.write(f"DIM_Service: {len(dim_service_rows)}건, DIM_Department: {len(dim_dept_rows)}건, DIM_OS: {len(dim_os_rows)}건\n")
    f.write(f"\nFACT_TACS_Matching 검증 시뮬레이션 (Asset 100건 기준):\n")
    for k, v in match_counts.items():
        f.write(f"  {k}: {v}건 ({v/len(all_asset_rows):.0%})\n")
    f.write(f"\n출력 위치: {OUT_DIR}\n")

print("done - see summary.txt")
