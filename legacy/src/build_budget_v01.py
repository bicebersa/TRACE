"""초기 산출물: 예산 대비 실적 리포트 v0.1 생성 (legacy/에 저장)."""
import datetime as dt, json, os, random, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as CL

OUTDIR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "예산실적리포트_v0.1")
os.makedirs(OUTDIR, exist_ok=True)
XLSX = os.path.join(OUTDIR, "예산실적리포트_v0.1.xlsx")

FONT = "맑은 고딕"
NUM = '#,##0;"△"#,##0;"-"'
PCT = '0.0%;"△"0.0%;"-"'
NAVY = PatternFill("solid", fgColor="1F3864")
GRAYH = PatternFill("solid", fgColor="808080")
SUB = PatternFill("solid", fgColor="D9E1F2")
TOT = PatternFill("solid", fgColor="F2F2F2")
YEL = PatternFill("solid", fgColor="FFFF00")
HELP = PatternFill("solid", fgColor="F7F7F7")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BLUE, BLACK, GREEN, GRAY = "0000FF", "000000", "008000", "7F7F7F"

def F(color=BLACK, bold=False, size=10):
    return Font(name=FONT, color=color, bold=bold, size=size)

def title(ws, t, s):
    ws["A1"] = t; ws["A1"].font = F(bold=True, size=14)
    ws["A2"] = s; ws["A2"].font = F(GRAY, size=9)

def header(ws, row, cols, start=1, fill=NAVY, h=30):
    for i, v in enumerate(cols):
        c = ws.cell(row=row, column=start + i, value=v)
        c.font = Font(name=FONT, color="FFFFFF", bold=True, size=10)
        c.fill = fill; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = h

def put(ws, r, c, v, color=BLACK, fmt=None, bold=False, fill=None, align=None, wrap=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = F(color, bold); cell.border = BOX
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    if align or wrap: cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    return cell

def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[CL(i)].width = x

# =========================================================== 예시 데이터 생성
rng = random.Random(7)
YEAR, MON = 2026, 9
ITEMS = {"B01": "여비교통비", "B02": "통신비", "B03": "소모품비", "B04": "도서인쇄비", "B05": "회의비", "B06": "업무추진비",
         "B07": "광고선전비", "B08": "교육훈련비", "B09": "지급수수료", "B10": "전산운영비", "B11": "임차료", "B12": "차량유지비"}
GLS = [  # G/L, 계정명, 항목, 항목 내 비중
    ("62101000", "여비교통비-국내", "B01", 0.85), ("62102000", "여비교통비-해외", "B01", 0.15),
    ("62201000", "통신비", "B02", 1.0), ("62301000", "소모품비", "B03", 1.0), ("62401000", "도서인쇄비", "B04", 1.0),
    ("62501000", "회의비", "B05", 1.0), ("62601000", "업무추진비", "B06", 1.0), ("62701000", "광고선전비", "B07", 1.0),
    ("62801000", "교육훈련비-사내", "B08", 0.4), ("62802000", "교육훈련비-외부위탁", "B08", 0.6),
    ("62901000", "지급수수료-컨설팅", "B09", 0.7), ("62902000", "지급수수료-기타", "B09", 0.3),
    ("63001000", "전산운영비-SW유지보수", "B10", 0.55), ("63002000", "전산운영비-클라우드", "B10", 0.45),
    ("63101000", "임차료", "B11", 1.0), ("63201000", "차량유지비", "B12", 1.0),
]
DEPTS = {"D100": "경영기획팀", "D200": "재무회계팀", "D300": "인사팀", "D400": "상품개발팀",
         "D500": "마케팅팀", "D600": "영업지원팀", "D700": "IT운영팀", "D800": "고객서비스팀"}
CCS = [  # 코스트센터, 이름, 부서, 부서 내 비중
    ("1101", "경영기획", "D100", 1.0), ("1201", "재무회계", "D200", 1.0), ("1301", "인사", "D300", 1.0),
    ("1401", "상품개발", "D400", 1.0), ("1501", "마케팅전략", "D500", 0.6), ("1502", "디지털마케팅", "D500", 0.4),
    ("1601", "영업지원-본사", "D600", 0.55), ("1602", "영업지원-지역본부", "D600", 0.45),
    ("1701", "IT운영", "D700", 1.0), ("1801", "고객서비스", "D800", 1.0),
]
BUDGET = {  # 연간 예산 (천원)
    "D100": {"B01": 36000, "B02": 9600, "B03": 6000, "B05": 12000, "B06": 48000, "B09": 240000},
    "D200": {"B01": 18000, "B02": 7200, "B03": 7200, "B05": 6000, "B09": 420000},
    "D300": {"B01": 24000, "B02": 8400, "B03": 9600, "B05": 9600, "B06": 24000, "B08": 480000},
    "D400": {"B01": 30000, "B02": 7200, "B03": 4800, "B04": 12000, "B05": 7200, "B09": 360000},
    "D500": {"B01": 42000, "B02": 12000, "B03": 9600, "B04": 84000, "B05": 14400, "B06": 36000, "B07": 1800000},
    "D600": {"B01": 96000, "B02": 26400, "B03": 14400, "B05": 18000, "B06": 42000, "B11": 216000, "B12": 96000},
    "D700": {"B01": 14400, "B02": 180000, "B03": 12000, "B05": 4800, "B10": 2400000},
    "D800": {"B01": 12000, "B02": 240000, "B03": 18000, "B04": 96000, "B05": 6000, "B09": 600000},
}
SEASON = {"B07": [0.6, 0.8, 1.4, 1.2, 1, 0.8, 0.7, 0.8, 1.4, 1.3, 1.0, 1.0],
          "B08": [0.5, 1, 1.3, 1.2, 1, 0.8, 0.6, 0.8, 1.2, 1.3, 1.2, 1.1]}
FACTOR = {("D500", "B07"): 1.28, ("D700", "B10"): 1.17, ("D400", "B09"): 0.48,
          ("D300", "B08"): 0.62, ("D600", "B12"): 1.22}

bud_month = {}  # (d,i) -> [12] 천원
for d, items in BUDGET.items():
    for i, ann in items.items():
        w = SEASON.get(i, [1] * 12)
        m = [round(ann * x / sum(w)) for x in w]
        m[11] += ann - sum(m)
        bud_month[(d, i)] = m

actual_rows = []  # (연도, 월, cc, ccname, gl, glname, 원)
def emit(year, month, d, i, amt_thousand):
    for cc, ccn, dd, cw in CCS:
        if dd != d: continue
        for gl, gln, ii, gw in GLS:
            if ii != i: continue
            won = round(amt_thousand * 1000 * cw * gw * rng.uniform(0.9, 1.1), -1)
            actual_rows.append((year, month, cc, ccn, gl, gln, won))

for (d, i), m in bud_month.items():
    f = FACTOR.get((d, i), rng.uniform(0.93, 1.03))
    g = rng.uniform(0.85, 1.0)
    for mo in range(1, 13):
        emit(YEAR - 1, mo, d, i, m[mo - 1] * g * rng.uniform(0.85, 1.15))
    for mo in range(1, MON + 1):
        emit(YEAR, mo, d, i, m[mo - 1] * f * rng.uniform(0.85, 1.15))
for mo, won in ((3, 4800000), (4, 3200000), (9, 5500000)):  # 무예산 집행 예시: 인사팀 채용광고
    actual_rows.append((YEAR, mo, "1301", "인사", "62701000", "광고선전비", won))
actual_rows.sort(key=lambda x: (x[0], x[1], x[2], x[4]))

AGG = sorted(set(bud_month) | {("D300", "B07")})
COMMENTS = {
    ("D500", "B07"): ("3분기 신상품 출시 캠페인(TV·디지털) 집행 앞당김. 4분기 집행분에서 조정 예정", "마케팅팀 김OO"),
    ("D700", "B10"): ("클라우드 사용량 증가, SW 유지보수 단가 인상(연 8%). 수정예산 반영 요청", "IT운영팀 이OO"),
    ("D400", "B09"): ("계리 검증 컨설팅 계약 지연 (11월 착수 예정)", "상품개발팀 박OO"),
    ("D300", "B08"): ("외부 위탁교육을 하반기에 집중 편성", "인사팀 최OO"),
    ("D300", "B07"): ("경력 채용 공고비. 예산 미편성 항목으로 전용 승인 진행 중", "인사팀 최OO"),
    ("D600", "B12"): ("지역본부 업무차량 정비비 일시 발생", "영업지원팀 정OO"),
}

# =========================================================== 레이아웃 상수
A_END = int(os.environ.get("A_END", 100000))  # 실적 집계 범위 끝 행
B_END = 3004            # 예산 집계 범위 끝 행
G_END = 506             # 집계 시트 범위 끝 행
A_FILL = 4 + len(actual_rows)   # 실적 보조수식 채운 마지막 행 (VBA가 데이터 길이에 맞춰 다시 채움)
B_FILL = 4 + len(bud_month)
S = {k: f"설정!$C${r}" for k, r in
     dict(co=5, yr=6, mon=7, ver=8, ub=9, ua=10, ur=11, th=12, mn=13, fb=14, fa=15).items()}
AR = lambda c: f"실적_원본!${c}$5:${c}${A_END}"
BR = lambda c: f"예산_원본!${c}$5:${c}${B_END}"
GR = lambda c: f"집계!${c}$7:${c}${G_END}"
MAP_ACT_TOP, MAP_BUD_TOP = 21, 32
ACT_STD = ["회계연도", "기간", "회사코드", "코스트센터", "코스트센터명", "G/L계정", "계정명", "금액"]
BUD_STD = ["연도", "예산버전", "부서코드", "부서명", "예산항목코드", "예산항목명"] + [f"{m}월" for m in range(1, 13)]

wb = Workbook()

# =========================================================== 안내
ws = wb.active; ws.title = "안내"
title(ws, "예산 대비 실적 리포트 v0.1", "Excel 2016 호환 · 예시 데이터는 가상 수치 · 단위는 설정 시트에서 변경")
blocks = [
    ("최초 설치 (1회)", [
        "1. 이 파일을 연 상태에서 Alt+F11 → [파일] → [파일 가져오기] → modBudget.bas 선택",
        "2. Alt+F8 → SetupButtons 실행 → 설정 시트에 버튼 6개가 생깁니다",
        "3. [다른 이름으로 저장] → 파일 형식 'Excel 매크로 사용 통합 문서 (*.xlsm)'",
        "4. 대시보드.html 을 이 엑셀 파일과 같은 폴더에 둡니다",
    ]),
    ("매월 사용 순서", [
        "1. 설정 시트에서 기준연도·기준월 변경",
        "2. [실적 불러오기] → 자체 화면에서 내려받은 파일 선택 (전년 실적 포함, 열 순서 무관)",
        "3. [예산 불러오기] → 예산 파일 선택 (예산이 바뀐 달에만)",
        "4. 자가진단 시트가 '전체 정상'인지 확인 (불러오기 후 자동으로 결과 팝업)",
        "5. 예외목록 확인 → 코멘트 시트에 원인·조치 입력 → 리포트·요약에 자동 반영",
        "6. [리포트 PDF] → [월마감 저장] → [대시보드 내보내기]",
    ]),
    ("매크로 없이 쓸 때", [
        "실적_원본 A~H열, 예산_원본 A~R열에 붙여넣고, 보조수식(회색 머리글 열)을 데이터 마지막 행까지 끌어서 채우면 됩니다.",
    ]),
    ("자체 화면 열 이름이 다를 때", [
        "설정 시트의 '열 매핑' 표에서 [원본 화면 열 이름] 칸을 실제 화면 헤더와 똑같이 고치세요. 불러오기가 그 이름으로 열을 찾습니다.",
        "합계 행, 빈 행(코스트센터·부서코드가 비어 있는 행)은 불러올 때 자동으로 제외됩니다.",
    ]),
    ("외부망으로 오류를 전달할 때", [
        "자가진단 시트의 [점검 항목]과 [상태]만 전달하세요. 금액·코드는 옮기지 않아도 원인을 찾을 수 있게 만들었습니다.",
    ]),
    ("셀 색상 규칙", [
        "파란 글씨 = 입력값 · 검정 = 수식 · 초록 = 다른 시트 참조 수식 · 노란 배경 = 매월 바꾸는 설정 · 회색 머리글 = 보조수식(수정 금지)",
    ]),
]
r = 4
for head, lines in blocks:
    ws.cell(row=r, column=1, value=head).font = F(bold=True, size=11); r += 1
    for ln in lines:
        ws.cell(row=r, column=1, value=ln).font = F(); r += 1
    r += 1
ws.cell(row=r, column=1, value="시트 구성").font = F(bold=True, size=11); r += 1
header(ws, r, ["시트", "역할"], h=20); r += 1
for s_, d_ in [("설정", "기준월·단위·경고 기준, 열 매핑, 매크로 버튼"), ("자가진단", "데이터 구조·매핑·대사 점검 (15개 항목)"),
               ("리포트", "인쇄용 1장 보고서 (PDF 출력 대상)"), ("요약", "부서별·예산항목별 요약, 부서별 보고 문구 초안"),
               ("예외목록", "경고 기준을 넘은 항목, 차이 금액 큰 순"), ("월별추이", "부서별 월 예산·실적·전년실적 (대시보드 원천)"),
               ("집계", "부서 × 예산항목 상세 (모든 계산의 중심)"), ("코멘트", "예외 항목 원인·조치 입력 (월별 누적)"),
               ("마감이력", "[월마감 저장] 시 집계 결과 보관"), ("예산_원본", "예산 파일 (부서 × 항목 × 12개월)"),
               ("실적_원본", "자체 화면 실적 (코스트센터 × G/L × 월)"), ("계정매핑", "G/L 계정 → 예산항목"),
               ("부서매핑", "코스트센터 → 부서")]:
    put(ws, r, 1, s_, bold=True); put(ws, r, 2, d_); r += 1
widths(ws, [16, 110])
ws.column_dimensions["A"].width = 16

# =========================================================== 설정
st = wb.create_sheet("설정")
title(st, "설정", "노란 칸만 바꾸세요 · 오른쪽 버튼은 SetupButtons 실행 후 생깁니다")
header(st, 4, ["항목", "값", "설명"], start=2, h=20)
srows = [
    (5, "회사명", "예시보험㈜", "리포트 제목에 표시"),
    (6, "기준연도", YEAR, "보고 대상 연도"),
    (7, "기준월", MON, "1~12. 누계는 1월~기준월"),
    (8, "적용 예산버전", "본예산", "예산_원본 B열(예산버전) 중 집계할 버전"),
    (9, "예산_원본 금액단위", "천원", "예산 파일에 적힌 금액 단위"),
    (10, "실적_원본 금액단위", "원", "자체 화면 금액 단위"),
    (11, "보고 단위", "천원", "모든 집계·리포트 표시 단위"),
    (12, "경고 기준: 집행률 차이", 0.10, "누계 집행률이 100% ± 이 값을 벗어나면 초과/과소"),
    (13, "경고 기준: 금액 하한 (보고단위)", 10000, "누계 차이가 이 금액 미만이면 집행률과 무관하게 정상"),
]
for rr, k, v, d in srows:
    put(st, rr, 2, k, bold=True, fill=SUB)
    put(st, rr, 3, v, BLUE, PCT if rr == 12 else (NUM if rr == 13 else None), fill=YEL, align="center")
    put(st, rr, 4, d, GRAY)
put(st, 14, 2, "(계산) 예산 환산계수", bold=True, fill=SUB)
put(st, 14, 3, "=INDEX($G$5:$G$7,MATCH(C9,$F$5:$F$7,0))/INDEX($G$5:$G$7,MATCH(C11,$F$5:$F$7,0))", fmt="0.######", align="center")
put(st, 14, 4, "예산_원본 금액 × 계수 = 보고단위", GRAY)
put(st, 15, 2, "(계산) 실적 환산계수", bold=True, fill=SUB)
put(st, 15, 3, "=INDEX($G$5:$G$7,MATCH(C10,$F$5:$F$7,0))/INDEX($G$5:$G$7,MATCH(C11,$F$5:$F$7,0))", fmt="0.######", align="center")
put(st, 15, 4, "실적_원본 금액 × 계수 = 보고단위", GRAY)
header(st, 4, ["단위", "원 환산"], start=6, h=20)
for i, (u, v) in enumerate([("원", 1), ("천원", 1000), ("백만원", 1000000)]):
    put(st, 5 + i, 6, u, BLUE, align="center"); put(st, 5 + i, 7, v, BLUE, "#,##0")
dvu = DataValidation(type="list", formula1="=$F$5:$F$7"); st.add_data_validation(dvu); dvu.add("C9:C11")
dvm = DataValidation(type="whole", operator="between", formula1="1", formula2="12"); st.add_data_validation(dvm); dvm.add("C7")
st["I4"] = "매크로 버튼"; st["I4"].font = F(bold=True)

st.cell(row=19, column=2, value="열 매핑 - 실적 (자체 화면)").font = F(bold=True, size=11)
header(st, 20, ["표준 열 (실적_원본)", "원본 화면 열 이름", "설명"], start=2, h=20)
act_desc = ["숫자·텍스트 모두 가능", "1~12 또는 001~012", "", "코스트센터 코드 (비어 있는 행은 제외)", "", "", "", "회사코드 통화, 비용 차변 +"]
for i, h in enumerate(ACT_STD):
    put(st, MAP_ACT_TOP + i, 2, f"{CL(i+1)}열 {h}", bold=True, fill=SUB)
    put(st, MAP_ACT_TOP + i, 3, h, BLUE, fill=YEL)
    put(st, MAP_ACT_TOP + i, 4, act_desc[i], GRAY)
st.cell(row=30, column=2, value="열 매핑 - 예산 파일").font = F(bold=True, size=11)
header(st, 31, ["표준 열 (예산_원본)", "원본 파일 열 이름", "설명"], start=2, h=20)
for i, h in enumerate(BUD_STD):
    put(st, MAP_BUD_TOP + i, 2, f"{CL(i+1)}열 {h}", bold=True, fill=SUB)
    put(st, MAP_BUD_TOP + i, 3, h, BLUE, fill=YEL)
    put(st, MAP_BUD_TOP + i, 4, "부서코드가 비어 있는 행은 제외" if i == 2 else "", GRAY)
widths(st, [2, 30, 18, 46, 3, 9, 12, 3, 24])

# =========================================================== 계정매핑 / 부서매핑
gm = wb.create_sheet("계정매핑")
title(gm, "계정 매핑 (G/L 계정 → 예산항목)", "코드 칸은 텍스트 형식 · 새 계정이 생기면 여기에 추가 (자가진단이 미매핑 계정을 알려줌)")
header(gm, 4, ["G/L 계정", "계정명", "예산항목코드", "예산항목명", "비고"])
for i, (gl, gln, it, _) in enumerate(GLS):
    for j, v in enumerate([gl, gln, it, ITEMS[it], None]):
        c = put(gm, 5 + i, j + 1, v, BLUE, align="center" if j in (0, 2) else None)
for row in gm.iter_rows(min_row=5, max_row=504, min_col=1, max_col=3):
    for c in row: c.number_format = "@"
gm.freeze_panes = "A5"; widths(gm, [12, 26, 13, 16, 30])

dm = wb.create_sheet("부서매핑")
title(dm, "부서 매핑 (코스트센터 → 부서)", "코드 칸은 텍스트 형식 · 조직개편 시 여기만 고치면 전체 반영")
header(dm, 4, ["코스트센터", "코스트센터명", "부서코드", "부서명", "비고"])
for i, (cc, ccn, d, _) in enumerate(CCS):
    for j, v in enumerate([cc, ccn, d, DEPTS[d], None]):
        put(dm, 5 + i, j + 1, v, BLUE, align="center" if j in (0, 2) else None)
for row in dm.iter_rows(min_row=5, max_row=504, min_col=1, max_col=3):
    for c in row: c.number_format = "@"
dm.freeze_panes = "A5"; widths(dm, [12, 20, 10, 14, 30])

# =========================================================== 실적_원본
ar = wb.create_sheet("실적_원본")
title(ar, "실적 원본 (자체 화면 추출)", "A~H열: 붙여넣기/불러오기 · J~Q열: 보조수식 (회색, 수정 금지) · 전년 실적도 함께 넣으세요")
header(ar, 4, ACT_STD)
header(ar, 4, ["부서코드", "예산항목코드", "연도(숫자)", "월(숫자)", "금액(보고단위)", "키", "예산유무", "집계포함"], start=10, fill=GRAYH)
for k, row in enumerate(actual_rows):
    r = 5 + k
    vals = [row[0], row[1], "1000", int(row[2]), row[3], int(row[4]), row[5], row[6]]
    for j, v in enumerate(vals):
        c = ar.cell(row=r, column=j + 1, value=v)
        c.font = F(BLUE)
        if j == 7: c.number_format = "#,##0"
    # 코스트센터·계정은 화면에서 숫자로 붙여넣어지는 경우를 가정해 숫자로 기록
for r in range(5, A_FILL + 1):
    fm = [
        f'=IF($A{r}="","",IFERROR(INDEX(부서매핑!$C$5:$C$504,MATCH($D{r}&"",부서매핑!$A$5:$A$504,0)),"미매핑"))',
        f'=IF($A{r}="","",IFERROR(INDEX(계정매핑!$C$5:$C$504,MATCH($F{r}&"",계정매핑!$A$5:$A$504,0)),"미매핑"))',
        f'=IF($A{r}="","",VALUE($A{r}&""))',
        f'=IF($A{r}="","",VALUE($B{r}&""))',
        f'=IF($A{r}="","",$H{r}*{S["fa"]})',
        f'=IF($A{r}="","",J{r}&"|"&K{r})',
        f'=IF($A{r}="","",IF(COUNTIFS({BR("T")},O{r},{BR("B")},{S["ver"]},{BR("U")},L{r})>0,"Y","N"))',
        f'=IF($A{r}="","",IF(COUNTIF({GR("E")},O{r})>0,"Y","N"))',
    ]
    for j, f_ in enumerate(fm):
        c = ar.cell(row=r, column=10 + j, value=f_)
        c.font = F(GREEN if j in (0, 1, 4, 6, 7) else BLACK)
        if j == 4: c.number_format = '#,##0.0;"△"#,##0.0;"-"'
ar.freeze_panes = "A5"; ar.auto_filter.ref = "A4:Q4"
widths(ar, [9, 6, 9, 11, 18, 11, 22, 15, 2, 9, 11, 9, 7, 14, 12, 8, 8])

# =========================================================== 예산_원본
br = wb.create_sheet("예산_원본")
title(br, "예산 원본", "A~R열: 붙여넣기/불러오기 (금액은 설정의 '예산_원본 금액단위') · T~Y열: 보조수식 (회색, 수정 금지)")
header(br, 4, BUD_STD[:6])
for m in range(1, 13):
    c = br.cell(row=4, column=6 + m, value=m)
    c.number_format = '0"월"'; c.font = Font(name=FONT, color="FFFFFF", bold=True); c.fill = NAVY; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center")
header(br, 4, ["비고"], start=19)
header(br, 4, ["키", "연도(숫자)", "연간예산\n(보고단위)", "당월예산\n(보고단위)", "누계예산\n(보고단위)", "집계포함"], start=20, fill=GRAYH)
for k, ((d, i), m) in enumerate(sorted(bud_month.items())):
    r = 5 + k
    for j, v in enumerate([YEAR, "본예산", d, DEPTS[d], i, ITEMS[i]] + m):
        c = br.cell(row=r, column=j + 1, value=v); c.font = F(BLUE)
        if j >= 6: c.number_format = "#,##0"
for r in range(5, B_FILL + 1):
    fm = [
        f'=IF($A{r}="","",$C{r}&"|"&$E{r})',
        f'=IF($A{r}="","",VALUE($A{r}&""))',
        f'=IF($A{r}="","",SUM($G{r}:$R{r})*{S["fb"]})',
        f'=IF($A{r}="","",INDEX($G{r}:$R{r},{S["mon"]})*{S["fb"]})',
        f'=IF($A{r}="","",SUMPRODUCT(($G$4:$R$4<={S["mon"]})*$G{r}:$R{r})*{S["fb"]})',
        f'=IF($A{r}="","",IF(COUNTIF({GR("E")},$T{r})>0,"Y","N"))',
    ]
    for j, f_ in enumerate(fm):
        c = br.cell(row=r, column=20 + j, value=f_)
        c.font = F(GREEN if j in (2, 3, 4, 5) else BLACK)
        if j in (2, 3, 4): c.number_format = NUM
br.freeze_panes = "G5"; br.auto_filter.ref = "A4:Y4"
widths(br, [7, 9, 9, 13, 11, 12] + [9] * 12 + [10, 13, 7, 12, 11, 11, 8])

# =========================================================== 집계
ag = wb.create_sheet("집계")
title(ag, "집계 (부서 × 예산항목)", "A·C열(파란색)만 입력 · 새 부서/항목은 매크로가 자동 추가 (매크로 없이 쓰면 직접 추가) · 차이 = 실적 − 예산")
ag["A3"] = f'="기준: "&{S["yr"]}&"년 "&{S["mon"]}&"월 · 예산버전: "&{S["ver"]}&" · 단위: "&{S["ur"]}'
ag["A3"].font = F(GREEN, size=9)
cols = ["부서코드", "부서명", "예산항목\n코드", "예산항목명", "키", "당월예산", "당월실적", "당월차이",
        "누계예산", "누계실적", "누계차이", "누계\n집행률", "연간예산", "잔여예산", "연간\n소진율",
        "전년동기\n누계실적", "전년비\n증감", "전년비\n증감률", "판정", "원인·조치 코멘트"]
header(ag, 6, cols)
header(ag, 6, ["예외점수", "초과점수"], start=21, fill=GRAYH)
put(ag, 4, 4, "합계", bold=True, fill=TOT)
for c, col in [(6, "F"), (7, "G"), (8, "H"), (9, "I"), (10, "J"), (11, "K"), (13, "M"), (14, "N"), (16, "P"), (17, "Q")]:
    put(ag, 4, c, f"=SUM({col}7:{col}{G_END})", fmt=NUM, bold=True, fill=TOT)
put(ag, 4, 12, "=IF(I4=0,0,J4/I4)", fmt=PCT, bold=True, fill=TOT)
put(ag, 4, 15, "=IF(M4=0,0,J4/M4)", fmt=PCT, bold=True, fill=TOT)
put(ag, 4, 18, "=IF(P4=0,0,J4/P4-1)", fmt=PCT, bold=True, fill=TOT)
AGG_PREFILL = 7 + len(AGG) + 40
akey, bkey = AR("O"), BR("T")
for idx in range(AGG_PREFILL - 7):
    r = 7 + idx
    d, i = AGG[idx] if idx < len(AGG) else (None, None)
    put(ag, r, 1, d, BLUE, align="center")
    put(ag, r, 3, i, BLUE, align="center")
    E = f'IF($A{r}="",""'
    bud = lambda col: f'=IF($A{r}="","",SUMIFS({BR(col)},{bkey},$E{r},{BR("B")},{S["ver"]},{BR("U")},{S["yr"]}))'
    fm = {
        2: (f'=IF($A{r}="","",IFERROR(INDEX(부서매핑!$D$5:$D$504,MATCH($A{r},부서매핑!$C$5:$C$504,0)),"부서매핑 없음"))', GREEN, None),
        4: (f'=IF($A{r}="","",IFERROR(INDEX(계정매핑!$D$5:$D$504,MATCH($C{r},계정매핑!$C$5:$C$504,0)),"항목매핑 없음"))', GREEN, None),
        5: (f'=IF($A{r}="","",$A{r}&"|"&$C{r})', BLACK, None),
        6: (bud("W"), GREEN, NUM),
        7: (f'=IF($A{r}="","",SUMIFS({AR("N")},{akey},$E{r},{AR("L")},{S["yr"]},{AR("M")},{S["mon"]}))', GREEN, NUM),
        8: (f'=IF($A{r}="","",G{r}-F{r})', BLACK, NUM),
        9: (bud("X"), GREEN, NUM),
        10: (f'=IF($A{r}="","",SUMIFS({AR("N")},{akey},$E{r},{AR("L")},{S["yr"]},{AR("M")},"<="&{S["mon"]}))', GREEN, NUM),
        11: (f'=IF($A{r}="","",J{r}-I{r})', BLACK, NUM),
        12: (f'=IF($A{r}="","",IF(I{r}=0,0,J{r}/I{r}))', BLACK, PCT),
        13: (bud("V"), GREEN, NUM),
        14: (f'=IF($A{r}="","",M{r}-J{r})', BLACK, NUM),
        15: (f'=IF($A{r}="","",IF(M{r}=0,0,J{r}/M{r}))', BLACK, PCT),
        16: (f'=IF($A{r}="","",SUMIFS({AR("N")},{akey},$E{r},{AR("L")},{S["yr"]}-1,{AR("M")},"<="&{S["mon"]}))', GREEN, NUM),
        17: (f'=IF($A{r}="","",J{r}-P{r})', BLACK, NUM),
        18: (f'=IF($A{r}="","",IF(P{r}=0,0,J{r}/P{r}-1))', BLACK, PCT),
        19: (f'=IF($A{r}="","",IF(AND(I{r}=0,J{r}<>0),"무예산집행",IF(ABS(K{r})<{S["mn"]},"정상",'
             f'IF(L{r}>1+{S["th"]},"초과",IF(L{r}<1-{S["th"]},"과소","정상")))))', GREEN, None),
        20: (f'=IF($A{r}="","",IFERROR(INDEX(코멘트!$E$5:$E$2004,MATCH({S["yr"]}&"|"&{S["mon"]}&"|"&$E{r},코멘트!$H$5:$H$2004,0)),""))', GREEN, None),
        21: (f'=IF($A{r}="","",IF(S{r}="정상",-1,ABS(K{r})+ROW()/1000000))', BLACK, "0.00"),
        22: (f'=IF($A{r}="",0,K{r}+ROW()/1000000)', BLACK, "0.00"),
    }
    for c, (f_, col, fmt) in fm.items():
        put(ag, r, c, f_, col, fmt, align="center" if c in (5, 19) else None)
ag.freeze_panes = "F7"; ag.auto_filter.ref = f"A6:T{AGG_PREFILL-1}"
widths(ag, [9, 13, 9, 13, 10, 11, 11, 11, 12, 12, 11, 8, 12, 12, 8, 12, 11, 8, 11, 44, 9, 9])
ag.column_dimensions["E"].hidden = True
rng_s = f"S7:S{G_END}"
ag.conditional_formatting.add(rng_s, CellIsRule(operator="equal", formula=['"초과"'], fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(color="9C0006", bold=True)))
ag.conditional_formatting.add(rng_s, CellIsRule(operator="equal", formula=['"과소"'], fill=PatternFill("solid", fgColor="FFE699")))
ag.conditional_formatting.add(rng_s, CellIsRule(operator="equal", formula=['"무예산집행"'], fill=PatternFill("solid", fgColor="D9C3E9"), font=Font(bold=True)))

# =========================================================== 코멘트
cm = wb.create_sheet("코멘트")
title(cm, "원인·조치 코멘트", "기준연도·기준월이 같은 코멘트가 집계·요약·예외목록·리포트에 표시됩니다 · 월별로 계속 쌓아 두세요")
header(cm, 4, ["연도", "월", "부서코드", "예산항목코드", "원인·조치 코멘트", "작성자", "작성일"])
header(cm, 4, ["키"], start=8, fill=GRAYH)
for k, ((d, i), (txt, who)) in enumerate(COMMENTS.items()):
    r = 5 + k
    for j, v in enumerate([YEAR, MON, d, i, txt, who, dt.date(2026, 10, 5)]):
        put(cm, r, j + 1, v, BLUE, "yyyy-mm-dd" if j == 6 else None, align="center" if j in (0, 1, 2, 3) else None)
for r in range(5, 2005):
    c = cm.cell(row=r, column=8, value=f'=IF($A{r}="","",$A{r}&"|"&$B{r}&"|"&$C{r}&"|"&$D{r})')
    c.font = F(GRAY)
cm.freeze_panes = "A5"; widths(cm, [7, 5, 9, 11, 60, 16, 11, 16])

# =========================================================== 예외목록
ex = wb.create_sheet("예외목록")
title(ex, "예외 항목 (누계 기준, 차이 금액 큰 순)", "판정 기준은 설정 시트 · 최대 30건 표시")
ex["A3"] = f'="기준: "&{S["yr"]}&"년 "&{S["mon"]}&"월 누계 · 단위: "&{S["ur"]}&" · 예외 "&COUNTIF({GR("U")},">=0")&"건"'
ex["A3"].font = F(GREEN, size=9)
header(ex, 5, ["순위", "부서명", "예산항목명", "판정", "누계예산", "누계실적", "누계차이", "집행률", "원인·조치 코멘트"])
header(ex, 5, ["행번호"], start=10, fill=GRAYH)
for k in range(1, 31):
    r = 5 + k
    put(ex, r, 1, k, align="center")
    put(ex, r, 10, f'=IF(IFERROR(LARGE({GR("U")},A{r}),-1)<0,"",MATCH(LARGE({GR("U")},A{r}),{GR("U")},0))', GRAY)
    for c, col, fmt in [(2, "B", None), (3, "D", None), (4, "S", None), (5, "I", NUM), (6, "J", NUM), (7, "K", NUM), (8, "L", PCT)]:
        put(ex, r, c, f'=IF($J{r}="","",INDEX({GR(col)},$J{r}))', GREEN, fmt, align="center" if c == 4 else None)
    put(ex, r, 9, f'=IF($J{r}="","",IF(INDEX({GR("T")},$J{r})="","(미입력 - 코멘트 시트에 작성)",INDEX({GR("T")},$J{r})))', GREEN)
for cf in (CellIsRule(operator="equal", formula=['"초과"'], fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(color="9C0006", bold=True)),
           CellIsRule(operator="equal", formula=['"과소"'], fill=PatternFill("solid", fgColor="FFE699")),
           CellIsRule(operator="equal", formula=['"무예산집행"'], fill=PatternFill("solid", fgColor="D9C3E9"), font=Font(bold=True))):
    ex.conditional_formatting.add("D6:D35", cf)
ex.freeze_panes = "A6"; widths(ex, [6, 13, 13, 11, 12, 12, 12, 8, 60, 7])

# =========================================================== 요약
sm = wb.create_sheet("요약")
title(sm, "요약 (부서별 · 예산항목별)", "부서·항목 목록은 파란 칸 · 새 부서가 생기면 행을 추가하세요 (자가진단이 누락을 알려줌)")
sm["A3"] = f'="기준: "&{S["yr"]}&"년 "&{S["mon"]}&"월 누계 · 단위: "&{S["ur"]}'
sm["A3"].font = F(GREEN, size=9)
sm["A5"] = "1. 부서별"; sm["A5"].font = F(bold=True, size=11)
header(sm, 6, ["부서코드", "부서명", "누계예산", "누계실적", "누계차이", "집행률", "연간예산", "잔여예산",
               "전년동기", "전년비", "예외건수", "주요 초과 항목", "보고 문구 초안"])
header(sm, 6, ["최대초과점수", "집계행"], start=14, fill=GRAYH)
D0 = 7
dl = list(DEPTS)
for k, d in enumerate(dl):
    r = D0 + k
    put(sm, r, 1, d, BLUE, align="center")
    put(sm, r, 2, f'=IFERROR(INDEX(부서매핑!$D$5:$D$504,MATCH($A{r},부서매핑!$C$5:$C$504,0)),"부서매핑 없음")', GREEN)
    for c, col in [(3, "I"), (4, "J"), (7, "M"), (9, "P")]:
        put(sm, r, c, f'=SUMIFS({GR(col)},{GR("A")},$A{r})', GREEN, NUM)
    put(sm, r, 5, f"=D{r}-C{r}", fmt=NUM)
    put(sm, r, 6, f"=IF(C{r}=0,0,D{r}/C{r})", fmt=PCT)
    put(sm, r, 8, f"=G{r}-D{r}", fmt=NUM)
    put(sm, r, 10, f"=IF(I{r}=0,0,D{r}/I{r}-1)", fmt=PCT)
    put(sm, r, 11, f'=COUNTIFS({GR("A")},$A{r},{GR("S")},"<>정상")', GREEN, NUM)
    put(sm, r, 14, f'=SUMPRODUCT(MAX(({GR("A")}=$A{r})*{GR("V")}))', GREEN, "0.00")
    put(sm, r, 15, f'=IF(N{r}<{S["mn"]},"",MATCH(N{r},{GR("V")},0))', GREEN)
    put(sm, r, 12, f'=IF(O{r}="","없음",INDEX({GR("D")},O{r})&" "&TEXT(INDEX({GR("K")},O{r}),"+#,##0"))', GREEN)
    put(sm, r, 13, (f'=B{r}&": 누계 집행률 "&TEXT(F{r},"0.0%")&" (예산 "&TEXT(C{r},"#,##0")&" / 실적 "&TEXT(D{r},"#,##0")&"), "'
                    f'&IF(I{r}=0,"전년 실적 없음","전년 동기 대비 "&TEXT(J{r},"+0.0%;-0.0%;0.0%"))'
                    f'&IF(O{r}="","",", 주요 초과 "&L{r})&IF(K{r}>0,", 예외 "&K{r}&"건","")&"."'))
D1 = D0 + len(dl) - 1
DT = D1 + 1
put(sm, DT, 1, "합계", bold=True, fill=TOT); put(sm, DT, 2, None, fill=TOT)
for c in (3, 4, 5, 7, 8, 9, 11):
    put(sm, DT, c, f"=SUM({CL(c)}{D0}:{CL(c)}{D1})", fmt=NUM, bold=True, fill=TOT)
put(sm, DT, 6, f"=IF(C{DT}=0,0,D{DT}/C{DT})", fmt=PCT, bold=True, fill=TOT)
put(sm, DT, 10, f"=IF(I{DT}=0,0,D{DT}/I{DT}-1)", fmt=PCT, bold=True, fill=TOT)
put(sm, DT, 12, None, fill=TOT)
put(sm, DT, 13, (f'="전사: 누계 집행률 "&TEXT(F{DT},"0.0%")&" (예산 "&TEXT(C{DT},"#,##0")&" / 실적 "&TEXT(D{DT},"#,##0")&"), 전년 동기 대비 "'
                 f'&TEXT(J{DT},"+0.0%;-0.0%;0.0%")&", 예외 "&K{DT}&"건."'), bold=True, fill=TOT)
I0 = DT + 4
sm.cell(row=I0 - 2, column=1, value="2. 예산항목별").font = F(bold=True, size=11)
header(sm, I0 - 1, ["항목코드", "예산항목명", "누계예산", "누계실적", "누계차이", "집행률", "연간예산", "잔여예산", "전년동기", "전년비", "예외건수"])
il = list(ITEMS)
for k, it in enumerate(il):
    r = I0 + k
    put(sm, r, 1, it, BLUE, align="center")
    put(sm, r, 2, f'=IFERROR(INDEX(계정매핑!$D$5:$D$504,MATCH($A{r},계정매핑!$C$5:$C$504,0)),"항목매핑 없음")', GREEN)
    for c, col in [(3, "I"), (4, "J"), (7, "M"), (9, "P")]:
        put(sm, r, c, f'=SUMIFS({GR(col)},{GR("C")},$A{r})', GREEN, NUM)
    put(sm, r, 5, f"=D{r}-C{r}", fmt=NUM)
    put(sm, r, 6, f"=IF(C{r}=0,0,D{r}/C{r})", fmt=PCT)
    put(sm, r, 8, f"=G{r}-D{r}", fmt=NUM)
    put(sm, r, 10, f"=IF(I{r}=0,0,D{r}/I{r}-1)", fmt=PCT)
    put(sm, r, 11, f'=COUNTIFS({GR("C")},$A{r},{GR("S")},"<>정상")', GREEN, NUM)
I1 = I0 + len(il) - 1
IT = I1 + 1
put(sm, IT, 1, "합계", bold=True, fill=TOT); put(sm, IT, 2, None, fill=TOT)
for c in (3, 4, 5, 7, 8, 9, 11):
    put(sm, IT, c, f"=SUM({CL(c)}{I0}:{CL(c)}{I1})", fmt=NUM, bold=True, fill=TOT)
put(sm, IT, 6, f"=IF(C{IT}=0,0,D{IT}/C{IT})", fmt=PCT, bold=True, fill=TOT)
put(sm, IT, 10, f"=IF(I{IT}=0,0,D{IT}/I{IT}-1)", fmt=PCT, bold=True, fill=TOT)
sm.freeze_panes = "C7"
widths(sm, [9, 13, 12, 12, 11, 8, 12, 12, 12, 8, 8, 22, 90, 11, 7])

# =========================================================== 월별추이
tr = wb.create_sheet("월별추이")
title(tr, "월별 추이 (예산 · 실적 · 전년실적)", "대시보드 원천 데이터 · 기준월 이후 실적은 빈칸")
tr["A3"] = f'="기준: "&{S["yr"]}&"년 · 단위: "&{S["ur"]}'; tr["A3"].font = F(GREEN, size=9)
header(tr, 5, ["부서코드", "부서명", "구분"])
for m in range(1, 13):
    c = tr.cell(row=5, column=3 + m, value=m)
    c.number_format = '0"월"'; c.font = Font(name=FONT, color="FFFFFF", bold=True); c.fill = NAVY; c.border = BOX
    c.alignment = Alignment(horizontal="center")
header(tr, 5, ["합계"], start=16)
T0 = 6
r = T0
for d in ["전체"] + dl:
    for kind in ("예산", "실적", "전년실적"):
        put(tr, r, 1, d, BLUE, align="center")
        put(tr, r, 2, "전사" if d == "전체" else f'=IFERROR(INDEX(부서매핑!$D$5:$D$504,MATCH($A{r},부서매핑!$C$5:$C$504,0)),"")', BLACK if d == "전체" else GREEN)
        put(tr, r, 3, kind, BLUE, align="center")
        for m in range(1, 13):
            col = CL(3 + m)
            dcrit_b = "" if d == "전체" else f',{BR("C")},$A{r}'
            dcrit_a = "" if d == "전체" else f',{AR("J")},$A{r}'
            if kind == "예산":
                f_ = f'=SUMIFS({BR(CL(6 + m))},{BR("B")},{S["ver"]},{BR("U")},{S["yr"]}{dcrit_b})*{S["fb"]}'
            elif kind == "실적":
                f_ = f'=IF({col}$5>{S["mon"]},"",SUMIFS({AR("N")},{AR("L")},{S["yr"]},{AR("M")},{col}$5{dcrit_a}))'
            else:
                f_ = f'=SUMIFS({AR("N")},{AR("L")},{S["yr"]}-1,{AR("M")},{col}$5{dcrit_a})'
            put(tr, r, 3 + m, f_, GREEN, NUM)
        put(tr, r, 16, f"=SUM(D{r}:O{r})", fmt=NUM, bold=True)
        if d == "전체":
            for c in range(1, 17): tr.cell(row=r, column=c).fill = SUB
        r += 1
T1 = r - 1
tr.freeze_panes = "D6"; widths(tr, [9, 13, 9] + [11] * 12 + [13])

# =========================================================== 마감이력
hs = wb.create_sheet("마감이력")
title(hs, "월마감 이력", "[월마감 저장] 매크로가 집계 결과를 값으로 보관합니다 · 같은 월을 다시 저장하면 덮어쓰기 확인")
header(hs, 4, ["연도", "월", "부서코드", "부서명", "예산항목코드", "예산항목명", "누계예산", "누계실적", "누계차이", "집행률", "판정", "코멘트", "저장일시"])
hs.freeze_panes = "A5"; widths(hs, [7, 5, 9, 13, 11, 13, 12, 12, 12, 8, 11, 50, 17])

# =========================================================== 리포트
rp = wb.create_sheet("리포트")
rp["A1"] = f'={S["co"]}&" 예산 집행 현황 ("&{S["yr"]}&"년 "&{S["mon"]}&"월 누계)"'
rp["A1"].font = F(bold=True, size=16)
rp["A2"] = f'="단위: "&{S["ur"]}&" · 예산버전: "&{S["ver"]}&" · 판정 기준: 집행률 ±"&TEXT({S["th"]},"0%")&", 차이 "&TEXT({S["mn"]},"#,##0")&" 이상"'
rp["A2"].font = F(GRAY, size=9)
kpis = [("연간예산", f"=요약!G{DT}", NUM), ("누계예산", f"=요약!C{DT}", NUM), ("누계실적", f"=요약!D{DT}", NUM),
        ("누계 집행률", f"=요약!F{DT}", PCT), ("전년 동기 대비", f"=요약!J{DT}", '+0.0%;-0.0%;0.0%'), ("예외 항목", f"=요약!K{DT}", '0"건"')]
for k, (lab, f_, fmt) in enumerate(kpis):
    c = rp.cell(row=4, column=1 + k, value=lab); c.font = F(GRAY, size=9); c.fill = TOT
    c.alignment = Alignment(horizontal="center"); c.border = BOX
    c = rp.cell(row=5, column=1 + k, value=f_); c.font = F(GREEN, bold=True, size=14); c.number_format = fmt
    c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX; c.fill = TOT
rp.row_dimensions[5].height = 30
rp["A7"] = "1. 부서별 집행 현황"; rp["A7"].font = F(bold=True, size=11)
header(rp, 8, ["부서", "누계예산", "누계실적", "누계차이", "집행률", "전년비", "예외", "요약"], h=20)
for k in range(len(dl)):
    r, sr = 9 + k, D0 + k
    for c, col, fmt in [(1, "B", None), (2, "C", NUM), (3, "D", NUM), (4, "E", NUM), (5, "F", PCT), (6, "J", '+0.0%;-0.0%;0.0%'), (7, "K", '0"건";;"-"')]:
        put(rp, r, c, f"=요약!{col}{sr}", GREEN, fmt, align="center" if c == 7 else None)
    put(rp, r, 8, f'=IF(요약!O{sr}="","",요약!L{sr})&IF(요약!K{sr}>0," / 예외 "&요약!K{sr}&"건","")', GREEN)
r = 9 + len(dl)
put(rp, r, 1, "합계", bold=True, fill=TOT)
for c, col, fmt in [(2, "C", NUM), (3, "D", NUM), (4, "E", NUM), (5, "F", PCT), (6, "J", '+0.0%;-0.0%;0.0%'), (7, "K", '0"건";;"-"')]:
    put(rp, r, c, f"=요약!{col}{DT}", GREEN, fmt, bold=True, fill=TOT, align="center" if c == 7 else None)
put(rp, r, 8, None, fill=TOT)
x0 = r + 2
rp.cell(row=x0, column=1, value="2. 예외 항목 TOP 10").font = F(bold=True, size=11)
header(rp, x0 + 1, ["부서", "예산항목", "판정", "누계차이", "집행률", "원인·조치"], h=20)
rp.merge_cells(start_row=x0 + 1, start_column=6, end_row=x0 + 1, end_column=8)
for k in range(10):
    r = x0 + 2 + k; er = 6 + k
    for c, col, fmt in [(1, "B", None), (2, "C", None), (3, "D", None), (4, "G", NUM), (5, "H", PCT)]:
        put(rp, r, c, f"=예외목록!{col}{er}", GREEN, fmt, align="center" if c == 3 else None)
    put(rp, r, 6, f"=예외목록!I{er}", GREEN, wrap=True)
    for c in (7, 8): put(rp, r, c, None)
    rp.merge_cells(start_row=r, start_column=6, end_row=r, end_column=8)
    rp.row_dimensions[r].height = 28
for cf in (CellIsRule(operator="equal", formula=['"초과"'], font=Font(color="9C0006", bold=True)),
           CellIsRule(operator="equal", formula=['"무예산집행"'], font=Font(color="7030A0", bold=True))):
    rp.conditional_formatting.add(f"C{x0+2}:C{x0+11}", cf)
ch = BarChart(); ch.type = "bar"; ch.style = 10
ch.title = "부서별 누계예산 vs 누계실적"; ch.y_axis.numFmt = "#,##0"
ch.add_data(Reference(sm, min_col=3, max_col=4, min_row=6, max_row=D1), titles_from_data=True)
ch.set_categories(Reference(sm, min_col=2, min_row=D0, max_row=D1))
ch.height = 9; ch.width = 26
rp.add_chart(ch, f"A{x0+13}")
widths(rp, [14, 14, 14, 13, 10, 10, 8, 52])
rp.page_setup.orientation = "landscape"; rp.page_setup.paperSize = rp.PAPERSIZE_A4
rp.page_setup.fitToWidth = 1; rp.page_setup.fitToHeight = 0
rp.sheet_properties.pageSetUpPr.fitToPage = True
rp.print_area = f"A1:H{x0+31}"

# =========================================================== 자가진단
ck = wb.create_sheet("자가진단")
title(ck, "자가진단", "불러오기 후 자동 점검 · 외부망으로 오류를 알릴 때는 [점검 항목]과 [상태]만 전달하세요")
header(ck, 6, ["No", "점검 항목", "결과", "상태", "조치 방법"], h=20)
amis = "&".join(f'IF(실적_원본!{CL(i+1)}4<>"{h}","{CL(i+1)} ","")' for i, h in enumerate(ACT_STD))
bmis = "&".join(f'IF(예산_원본!{CL(i+1)}4<>"{h}","{CL(i+1)} ","")' for i, h in enumerate(BUD_STD[:6]))
checks = [
    ("설정값 유효성", f'=IF(AND(ISNUMBER({S["mon"]}),{S["mon"]}>=1,{S["mon"]}<=12,ISNUMBER({S["fb"]}),ISNUMBER({S["fa"]})),"정상","기준월 또는 단위 오류")',
     '=IF(C{r}="정상","OK","확인필요")', "설정 시트의 기준월(1~12)과 단위(원/천원/백만원) 확인"),
    ("실적_원본 머리글", f'=IF({amis}="","일치","불일치 열: "&{amis})', '=IF(C{r}="일치","OK","확인필요")', "실적_원본 4행 머리글을 표준 이름으로 되돌리기"),
    ("예산_원본 머리글", f'=IF(AND({bmis}="",예산_원본!G4=1,예산_원본!R4=12),"일치","불일치 열: "&{bmis}&IF(예산_원본!G4<>1,"G(1월) ",""))',
     '=IF(C{r}="일치","OK","확인필요")', "예산_원본 4행 머리글을 표준 이름으로 되돌리기"),
    ("실적 행 수 / 보조수식 범위", f'=COUNTA(실적_원본!A5:A{A_END})&"행 / 보조수식 "&COUNTA(실적_원본!O5:O{A_END})&"행"',
     f'=IF(AND(COUNTA(실적_원본!A5:A{A_END})<={A_END-4},COUNTA(실적_원본!O5:O{A_END})>=COUNTA(실적_원본!A5:A{A_END})),"OK","확인필요")',
     "보조수식(J~Q열)을 데이터 마지막 행까지 채우기 ([보조수식 다시 채우기] 버튼)"),
    ("예산 행 수 / 보조수식 범위", f'=COUNTA(예산_원본!A5:A{B_END})&"행 / 보조수식 "&COUNTA(예산_원본!T5:T{B_END})&"행"',
     f'=IF(COUNTA(예산_원본!T5:T{B_END})>=COUNTA(예산_원본!A5:A{B_END}),"OK","확인필요")', "보조수식(T~Y열)을 데이터 마지막 행까지 채우기"),
    ("기준월 실적 존재", f'=COUNTIFS({AR("L")},{S["yr"]},{AR("M")},{S["mon"]})&"건"',
     f'=IF(COUNTIFS({AR("L")},{S["yr"]},{AR("M")},{S["mon"]})>0,"OK","확인필요")', "기준월 실적이 포함된 파일인지, 기간 열 형식 확인"),
    ("전년 실적 존재", f'=COUNTIFS({AR("L")},{S["yr"]}-1)&"건"',
     f'=IF(COUNTIFS({AR("L")},{S["yr"]}-1)>0,"OK","참고")', "전년 대비가 필요하면 전년 실적도 함께 불러오기"),
    ("미매핑 코스트센터", f'=COUNTIF({AR("J")},"미매핑")&"건"&IF(COUNTIF({AR("J")},"미매핑")>0," (예: "&INDEX({AR("D")},MATCH("미매핑",{AR("J")},0))&")","")',
     f'=IF(COUNTIF({AR("J")},"미매핑")=0,"OK","확인필요")', "부서매핑 시트에 해당 코스트센터 추가"),
    ("미매핑 G/L 계정", f'=COUNTIF({AR("K")},"미매핑")&"건"&IF(COUNTIF({AR("K")},"미매핑")>0," (예: "&INDEX({AR("F")},MATCH("미매핑",{AR("K")},0))&")","")',
     f'=IF(COUNTIF({AR("K")},"미매핑")=0,"OK","확인필요")', "계정매핑 시트에 해당 계정 추가 (예산 통제 대상이 아니면 예산항목코드에 '제외' 입력)"),
    ("적용버전 예산 존재", f'=COUNTIFS({BR("B")},{S["ver"]},{BR("U")},{S["yr"]})&"행"',
     f'=IF(COUNTIFS({BR("B")},{S["ver"]},{BR("U")},{S["yr"]})>0,"OK","확인필요")', "설정의 '적용 예산버전'이 예산_원본 B열 값과 같은지 확인"),
    ("집계표에 없는 예산 조합", f'=COUNTIFS({BR("Y")},"N",{BR("B")},{S["ver"]},{BR("U")},{S["yr"]})&"건"',
     f'=IF(COUNTIFS({BR("Y")},"N",{BR("B")},{S["ver"]},{BR("U")},{S["yr"]})=0,"OK","확인필요")', "집계 시트에 부서코드·항목코드 행 추가 (매크로 불러오기 시 자동)"),
    ("집계표에 없는 실적 조합 (당해)", f'=COUNTIFS({AR("Q")},"N",{AR("L")},{S["yr"]})&"건"',
     f'=IF(COUNTIFS({AR("Q")},"N",{AR("L")},{S["yr"]})=0,"OK","확인필요")', "미매핑이 아니라면 집계 시트에 행 추가 (매크로 불러오기 시 자동)"),
    ("무예산 집행 (예산 없는 실적)", f'=COUNTIFS({AR("P")},"N",{AR("L")},{S["yr"]},{AR("M")},"<="&{S["mon"]})&"건"',
     f'=IF(COUNTIFS({AR("P")},"N",{AR("L")},{S["yr"]},{AR("M")},"<="&{S["mon"]})=0,"OK","참고")', "업무상 예외 - 예외목록에서 전용·추가예산 여부 확인"),
    ("실적 총액 대사 (원본 ↔ 집계)", f'="차이 "&TEXT(SUMIFS({AR("N")},{AR("L")},{S["yr"]},{AR("M")},"<="&{S["mon"]})-집계!J4,"#,##0")',
     f'=IF(ABS(SUMIFS({AR("N")},{AR("L")},{S["yr"]},{AR("M")},"<="&{S["mon"]})-집계!J4)<1,"OK","확인필요")', "미매핑·집계 누락 항목을 먼저 해결하면 대부분 0이 됩니다"),
    ("예산 총액 대사 (원본 ↔ 집계)", f'="차이 "&TEXT(SUMIFS({BR("V")},{BR("B")},{S["ver"]},{BR("U")},{S["yr"]})-집계!M4,"#,##0")',
     f'=IF(ABS(SUMIFS({BR("V")},{BR("B")},{S["ver"]},{BR("U")},{S["yr"]})-집계!M4)<1,"OK","확인필요")', "집계표에 없는 예산 조합 해결"),
    ("매핑 코드 중복", '="계정 "&SUMPRODUCT((계정매핑!$A$5:$A$504<>"")*(COUNTIF(계정매핑!$A$5:$A$504,계정매핑!$A$5:$A$504)>1))&"건 / 코스트센터 "&SUMPRODUCT((부서매핑!$A$5:$A$504<>"")*(COUNTIF(부서매핑!$A$5:$A$504,부서매핑!$A$5:$A$504)>1))&"건"',
     '=IF(SUMPRODUCT((계정매핑!$A$5:$A$504<>"")*(COUNTIF(계정매핑!$A$5:$A$504,계정매핑!$A$5:$A$504)>1))+SUMPRODUCT((부서매핑!$A$5:$A$504<>"")*(COUNTIF(부서매핑!$A$5:$A$504,부서매핑!$A$5:$A$504)>1))=0,"OK","확인필요")',
     "매핑 시트에서 같은 코드가 두 번 있는 행 정리"),
    ("요약 시트 부서 누락", f'=SUMPRODUCT(({GR("A")}<>"")*(COUNTIF(요약!$A${D0}:$A${D1},{GR("A")})=0))&"행"',
     f'=IF(SUMPRODUCT(({GR("A")}<>"")*(COUNTIF(요약!$A${D0}:$A${D1},{GR("A")})=0))=0,"OK","확인필요")', "요약 시트 부서별 표에 부서코드 행 추가"),
]
C0 = 7
for k, (nm, res, stt, how) in enumerate(checks):
    r = C0 + k
    put(ck, r, 1, k + 1, align="center")
    put(ck, r, 2, nm, bold=True)
    put(ck, r, 3, res, GREEN)
    put(ck, r, 4, stt.format(r=r), align="center", bold=True)
    put(ck, r, 5, how, GRAY)
C1 = C0 + len(checks) - 1
ck["A4"] = "종합"; ck["A4"].font = F(bold=True, size=12)
ck["B4"] = f'=IF(COUNTIF(D{C0}:D{C1},"확인필요")=0,"전체 정상"&IF(COUNTIF(D{C0}:D{C1},"참고")>0," (참고 "&COUNTIF(D{C0}:D{C1},"참고")&"건)",""),"확인필요 "&COUNTIF(D{C0}:D{C1},"확인필요")&"건")'
ck["B4"].font = F(bold=True, size=12)
okf = PatternFill("solid", fgColor="C6EFCE"); ngf = PatternFill("solid", fgColor="FFC7CE"); inf = PatternFill("solid", fgColor="FFEB9C")
ck.conditional_formatting.add(f"D{C0}:D{C1}", CellIsRule(operator="equal", formula=['"OK"'], fill=okf))
ck.conditional_formatting.add(f"D{C0}:D{C1}", CellIsRule(operator="equal", formula=['"확인필요"'], fill=ngf))
ck.conditional_formatting.add(f"D{C0}:D{C1}", CellIsRule(operator="equal", formula=['"참고"'], fill=inf))
ck.conditional_formatting.add("B4", FormulaRule(formula=['LEFT($B$4,4)="전체 정상"'], fill=okf))
ck.conditional_formatting.add("B4", FormulaRule(formula=['LEFT($B$4,4)="확인필요"'], fill=ngf))
widths(ck, [6, 30, 34, 10, 70])

# 시트 순서 · 공통
order = ["안내", "설정", "자가진단", "리포트", "요약", "예외목록", "월별추이", "집계", "코멘트", "마감이력", "예산_원본", "실적_원본", "계정매핑", "부서매핑"]
wb._sheets = [wb[n] for n in order]
for w in wb.worksheets:
    w.sheet_view.showGridLines = False
for n, colr in [("설정", "FFC000"), ("자가진단", "70AD47"), ("리포트", "1F3864"), ("예산_원본", "5B9BD5"), ("실적_원본", "5B9BD5")]:
    wb[n].sheet_properties.tabColor = colr
wb.active = 1
wb.calculation.fullCalcOnLoad = True
wb.save(XLSX)

LAYOUT = dict(A_END=A_END, B_END=B_END, G_END=G_END, MAP_ACT_TOP=MAP_ACT_TOP, MAP_BUD_TOP=MAP_BUD_TOP,
              C0=C0, C1=C1, T0=T0, T1=T1, AGG_PREFILL=AGG_PREFILL)
json.dump(LAYOUT, open(os.path.join(os.path.dirname(__file__), "layout.json"), "w"))
print("saved", XLSX, "actual rows", len(actual_rows), "budget rows", len(bud_month), "agg", len(AGG), LAYOUT)

# =========================================================== 기대값 (검증 · 예시 data.js용)
FA = 1 / 1000
def cum(lst, m): return sum(lst[:m])
exp = []
for d, i in AGG:
    bm = bud_month.get((d, i), [0] * 12)
    acts = [x for x in actual_rows if next(c[2] for c in CCS if c[0] == x[2]) == d and next(g[2] for g in GLS if g[0] == x[4]) == i]
    a = lambda y, cond: sum(x[6] for x in acts if x[0] == y and cond(x[1])) * FA
    exp.append(dict(d=d, i=i, mb=bm[MON - 1], ma=a(YEAR, lambda m: m == MON), cb=cum(bm, MON), ca=a(YEAR, lambda m: m <= MON),
                    ann=sum(bm), pc=a(YEAR - 1, lambda m: m <= MON)))
trend = []
for d in ["전체"] + dl:
    def inD(x): return d == "전체" or next(c[2] for c in CCS if c[0] == x[2]) == d
    b = [sum(v[m] for (dd, ii), v in bud_month.items() if d == "전체" or dd == d) for m in range(12)]
    a = [sum(x[6] for x in actual_rows if x[0] == YEAR and x[1] == m + 1 and inD(x)) * FA if m < MON else None for m in range(12)]
    p = [sum(x[6] for x in actual_rows if x[0] == YEAR - 1 and x[1] == m + 1 and inD(x)) * FA for m in range(12)]
    trend.append((d, b, a, p))
json.dump(dict(exp=exp, trend=trend), open(os.path.join(os.path.dirname(__file__), "expected.json"), "w"), ensure_ascii=False)
