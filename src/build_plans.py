"""연도별 예산 계획(배정 예산 원천) 예시 파일 생성.

열: 부서코드, 계정코드, 1월~12월, 총계, 신규여부, 청라여부, 항목, 산출근거 (항목 단위 행)
- 2026년 계획: 부서 × 계정 월별 합계가 예시 예산판 배정(tool_data.json bud_gl)과 정확히 같다.
- 2027년 계획: 단가·물량 변화, 신규 항목, 청라 관련 항목, 폐지 항목을 섞은 가상 계획.
결과: samples/예산계획_2026.xlsx, samples/예산계획_2027.xlsx, src/plans.json (대시보드 예시 데이터용)

사용: python src/build_plans.py   (build_trace.py 다음에 실행)
"""
import json, os, random
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
D = json.load(open(os.path.join(HERE, "tool_data.json"), encoding="utf-8"))
BUD = {tuple(k.split("|")): v for k, v in D["bud_gl"].items()}   # (부서, 계정) → 12개월 원
GLN = {}
for e in D["E"]: GLN.setdefault(e["gl"], e["gl"])
ITEMS = {  # 계정 → [(항목, 비중, 산출근거)]
    "62101000": [("본사-지역 업무 출장", .7, "월 {n}회 × 회당 {u}천원"), ("교육·세미나 참석 출장", .3, "연 {n}회 × 회당 {u}천원")],
    "62102000": [("해외 연수·컨퍼런스", 1.0, "연 {n}회 × 회당 {u}천원")],
    "62201000": [("유선·인터넷 회선", .6, "회선 {n}개 × 월 {u}천원"), ("업무용 휴대폰", .4, "{n}대 × 월 {u}천원")],
    "62301000": [("사무용품", .7, "인원 {n}명 × 월 {u}천원"), ("토너·용지", .3, "월 {n}박스 × {u}천원")],
    "62401000": [("상품 안내장 인쇄", .8, "연 {n}만부 × 부당 {u}원"), ("업무 도서 구입", .2, "연 {n}권 × {u}천원")],
    "62501000": [("정기 회의 다과", 1.0, "월 {n}회 × 회당 {u}천원")],
    "62601000": [("대외 업무추진", 1.0, "월 {n}건 × 건당 {u}천원")],
    "62701000": [("TV·디지털 광고", .75, "분기 캠페인 {n}회 × {u}백만원"), ("브랜드 협찬", .25, "연 {n}건 × {u}백만원")],
    "62801000": [("사내 집합교육", 1.0, "연 {n}과정 × 과정당 {u}천원")],
    "62802000": [("외부 위탁교육", 1.0, "연 {n}명 × 1인당 {u}천원")],
    "62901000": [("경영·회계 자문", 1.0, "연 {n}건 × 건당 {u}백만원")],
    "62902000": [("업무 위탁 수수료", 1.0, "월 {n}건 × 건당 {u}천원")],
    "63001000": [("SW 유지보수", 1.0, "솔루션 {n}종 × 연 {u}백만원")],
    "63002000": [("클라우드 사용료", 1.0, "월 {n}vCPU × {u}천원")],
    "63101000": [("본사 사무실 임차", 1.0, "{n}평 × 월 {u}천원")],
    "63201000": [("업무 차량 유지", 1.0, "차량 {n}대 × 월 {u}천원")],
    "63501000": [("GA 모집수수료", .6, "예상 신계약 {n}건 × 건당 {u}천원"), ("전속 모집수수료", .4, "예상 신계약 {n}건 × 건당 {u}천원")],
    "63502000": [("유지수수료", 1.0, "보유계약 {n}만건 × {u}원")],
    "63601000": [("손해사정 위탁", 1.0, "월 {n}건 × 건당 {u}천원")],
}
rng = random.Random(77)


def basis(tpl, total):
    n = rng.randint(2, 40); u = max(1, round(total / 1000 / max(n, 1) / 12))
    return tpl.format(n=n, u=f"{u:,}")


def split_month(v, ws):   # 월 금액을 항목 비중으로 나누고 끝전은 마지막 항목에
    parts = [round(v * w, -3) for w in ws[:-1]]
    return parts + [v - sum(parts)]


def plan2026():
    rows = []
    for (d, gl), months in sorted(BUD.items()):
        its = ITEMS.get(gl, [("기타", 1.0, "")])
        per = [split_month(v, [w for _, w, _ in its]) for v in months]
        for j, (nm, w, tpl) in enumerate(its):
            ms = [per[m][j] for m in range(12)]
            rows.append(dict(dept=d, gl=gl, months=ms, new="N", cheongna="N", item=nm, basis=basis(tpl, sum(ms)) if tpl else ""))
    return rows


def plan2027(base):
    rows = []
    for r in base:
        if r["dept"] == "D400" and r["item"] == "업무 도서 구입": continue        # 폐지
        f = rng.uniform(0.97, 1.08)
        if r["gl"] == "63002000": f = 1.22                                         # 클라우드 단가 인상 반영
        if r["gl"] == "62701000" and r["item"] == "TV·디지털 광고": f = 1.15       # 신상품 캠페인 확대
        if r["gl"] == "62802000": f = 0.85                                         # 위탁교육 축소
        ms = [round(v * f, -3) for v in r["months"]]
        rows.append(dict(r, months=ms, basis=r["basis"] + (" (단가 인상 반영)" if f > 1.1 else (" (축소)" if f < 0.9 else ""))))
    add = [("D600", "63101000", "청라 신사옥 임차 (하반기 입주)", "N", "Y", [0] * 6 + [38_000_000] * 6, "1,200평 × 월 31.7천원, 7월 입주"),
           ("D700", "63001000", "청라 IDC 이전 SW 재계약", "Y", "Y", [0] * 5 + [45_000_000] * 7, "솔루션 6종 이전 재계약, 6월부터"),
           ("D600", "62101000", "청라 이전 준비 출장", "Y", "Y", [3_000_000] * 6 + [0] * 6, "월 12회 × 회당 250천원, 상반기"),
           ("D500", "62701000", "신상품 출시 캠페인", "Y", "N", [0, 0, 120_000_000, 80_000_000] + [0] * 8, "4월 출시 TV·디지털 캠페인 2회"),
           ("D900", "63501000", "신채널 제휴 모집수수료", "Y", "N", [0] * 3 + [60_000_000] * 9, "제휴 신계약 월 400건 × 150천원"),
           ("D300", "62801000", "디지털 역량 사내교육", "Y", "N", [4_000_000] * 12, "연 8과정 × 과정당 6,000천원")]
    for d, gl, nm, new, cn, ms, b in add:
        rows.append(dict(dept=d, gl=gl, months=ms, new=new, cheongna=cn, item=nm, basis=b))
    return rows


HDR = ["부서코드", "계정코드"] + [f"{m}월" for m in range(1, 13)] + ["총계", "신규여부", "청라여부", "항목", "산출근거"]


def write(rows, year, path):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = f"{year}년 예산계획"
    thin = Side(style="thin", color="BFBFBF")
    for j, h in enumerate(HDR, 1):
        c = ws.cell(row=1, column=j, value=h); c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9D9D9")
        c.alignment = Alignment(horizontal="center"); c.border = Border(bottom=thin)
    for i, r in enumerate(rows, 2):
        vals = [r["dept"], r["gl"]] + r["months"] + [sum(r["months"]), r["new"], r["cheongna"], r["item"], r["basis"]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=i, column=j, value=v)
            if 3 <= j <= 15: c.number_format = "#,##0"
            if j in (1, 2): c.number_format = "@"
    for j, w in enumerate([9, 11] + [12] * 13 + [8, 8, 26, 40], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = "C2"
    wb.save(path)


def main():
    p26 = plan2026(); p27 = plan2027(p26)
    for (d, gl), months in BUD.items():   # 2026 계획 = 예산판 배정 (월별)
        got = [sum(r["months"][m] for r in p26 if (r["dept"], r["gl"]) == (d, gl)) for m in range(12)]
        assert all(abs(a - b) < 1 for a, b in zip(got, months)), (d, gl)
    os.makedirs(os.path.join(ROOT, "samples"), exist_ok=True)
    write(p26, 2026, os.path.join(ROOT, "samples", "예산계획_2026.xlsx"))
    write(p27, 2027, os.path.join(ROOT, "samples", "예산계획_2027.xlsx"))
    js = lambda rows: [dict(dept=r["dept"], gl=r["gl"], m=[x / 1e6 for x in r["months"]], total=sum(r["months"]) / 1e6, new=r["new"], cheongna=r["cheongna"],
                            item=r["item"], basis=r["basis"]) for r in rows]
    json.dump(dict(base=dict(year=2026, file="예산계획_2026.xlsx", reserve=1500.0, rows=js(p26)), cmp=dict(year=2027, file="예산계획_2027.xlsx", reserve=1800.0, rows=js(p27))),   # 예비비(백만원, 예시) → 편성 = 배정 + 예비비
              open(os.path.join(HERE, "plans.json"), "w", encoding="utf-8"), ensure_ascii=False)
    print("saved samples/예산계획_2026.xlsx", len(p26), "· 2027", len(p27), "· 2026 합계", round(sum(sum(r["months"]) for r in p26) / 1e6, 1),
          "· 2027 합계", round(sum(sum(r["months"]) for r in p27) / 1e6, 1))


if __name__ == "__main__":
    main()
