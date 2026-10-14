"""엑셀 수식 로직을 그대로 옮겨 예시 data.js를 만든다 (VBA ExportDashboard 출력과 같은 구조)."""
import json, math, os, statistics as stt, sys, datetime as dt
HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(HERE, "tool_data.json")))
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data.js")
E = D["E"]; BP = D["BP"]; KEYS = [tuple(k) for k in D["KEYS"]]
Y, M = 2026, 9; P = Y * 12 + M; QS = Y * 12 + 7
DEPTS = {"D100": "경영기획팀", "D200": "재무회계팀", "D300": "인사팀", "D400": "상품개발팀", "D500": "마케팅팀", "D600": "영업지원팀", "D700": "IT운영팀", "D800": "고객서비스팀", "D900": "영업채널팀", "D950": "자산운용팀"}
GLN = {"62101000": "여비교통비-국내", "62102000": "여비교통비-해외", "62201000": "통신비", "62301000": "소모품비", "62401000": "도서인쇄비", "62501000": "회의비",
       "62601000": "업무추진비", "62701000": "광고선전비", "62801000": "교육훈련비-사내", "62802000": "교육훈련비-외부위탁", "62901000": "지급수수료-컨설팅",
       "62902000": "지급수수료-기타", "63001000": "전산운영비-SW유지보수", "63002000": "전산운영비-클라우드", "63101000": "임차료", "63201000": "차량유지비", "63501000": "모집수수료", "63502000": "유지수수료"}
CC2D = {"1101": "D100", "1201": "D200", "1301": "D300", "1401": "D400", "1501": "D500", "1502": "D500", "1601": "D600", "1602": "D600", "1701": "D700", "1801": "D800", "1901": "D900", "1902": "D900", "1951": "D950"}
KW = [("연간 일괄 계상", ["연간", "일시", "선급", "1년", "반기"], "연간·반기 계약을 한 번에 계상했을 가능성. 월할 안분 대상인지 확인"),
      ("캠페인·행사", ["캠페인", "행사", "이벤트", "프로모션", "박람회"], "일회성 캠페인·행사 비용일 가능성"),
      ("정정·취소", ["정정", "취소", "역분개", "오류", "환입"], "정정 전표 영향. 순액 기준으로 다시 확인"),
      ("이관·재분류", ["재분류", "이관", "대체", "귀속변경", ""], "다른 부서·계정에서 넘어온 금액일 가능성"),
      ("추정 계상", ["미지급", "추정", "선계상", "", ""], "월말 추정 계상분. 다음 달 역분개 여부 확인"),
      ("소급·정산", ["소급", "정산", "추가청구", "차액", ""], "과거분 정산·소급 지급일 가능성"),
      ("판매시책·추가수수료", ["시책", "추가 수수료", "인센티브", "특별수당", ""], "판매 시책비·추가 수수료. 신계약비 경험차이로 GMM CSM에 영향 가능")]
COMMENTS = {"D500|62701000": "3분기 신상품 TV캠페인 집행을 앞당김. 4분기 집행분에서 조정 예정",
            "D200|62901000": "8월 경영기획팀에 잘못 귀속된 회계자문료를 재분류. 실제 증가 아님"}
for i, e in enumerate(E):
    e["row"] = 5 + i; e["p"] = e["y"] * 12 + e["m"]; e["key"] = CC2D[e["cc"]] + "|" + e["gl"]; e["v"] = e["amt"] / 1e6
def tot(key, p0, p1): return sum(e["v"] for e in E if e["key"] == key and p0 <= e["p"] <= p1)
def rate(a, b): return (1 if a > 0 else 0) if b <= 0 else a / b - 1
def t1(x): return f"{x:,.1f}"
def pct(x): return f"{math.floor(abs(x) * 100 + 0.5) * (1 if x >= 0 else -1):.0f}%"
def spct(x):
    s = pct(x); return s if s.startswith("-") or s == "0%" else "+" + s
def plab(p): return f"{(p - 1) // 12}-{(p - 1) % 12 + 1:02d}"

# ---------------- 변동분석
VA = {}
for d, gl in KEYS:
    k = f"{d}|{gl}"
    mon = {p: tot(k, p, p) for p in range(P - 23, P + 1)}
    F, G, I = mon[P], mon[P - 1], mon[P - 12]
    N = max(F - G, F - I)
    prev12 = [mon[p] for p in range(P - 12, P)]
    try: z = (F - stt.mean(prev12)) / stt.stdev(prev12)
    except ZeroDivisionError: z = 0
    avail = sum((r[6] + r[7]) / 1e6 for r in BP if f"{r[1]}|{r[3]}" == k)
    Q = tot(k, Y * 12 + 1, P); R = avail - Q
    Sa = (mon[P - 2] + mon[P - 1] + mon[P]) / 3 * (12 - M)
    py = tot(k, (Y - 1) * 12 + 1, P - 12); pa = tot(k, (Y - 1) * 12 + 1, (Y - 1) * 12 + 12)
    Tb = Q / (py / pa) - Q if py and pa else Sa
    U = Q + max(Sa, Tb); V = U / avail if avail else 0
    W = "무예산" if avail == 0 and Q > 0 else "-" if avail == 0 else "부족 예상" if V > 1 else "주의" if V > 0.9 else "여유"
    X = 0 if avail == 0 else max(0, U - avail)
    Ym = "" if W != "부족 예상" else ("이미 초과" if R <= 0 else f"{min(12, M + math.ceil(R / (max(Sa, Tb) / (12 - M))))}월")
    curE = [e for e in E if e["key"] == k and e["p"] == P]
    AC = (max([0] + [e["v"] for e in curE]) / N) if N > 0 else None
    AD = sum(e["v"] for e in curE if e["typ"] == "SA") - sum(e["v"] for e in E if e["key"] == k and e["p"] == P - 1 and e["typ"] == "SA")
    AE = sum(1 for e in curE if e["v"] < 0)
    sig = ""
    if AC is not None and AC >= 0.5 and N >= 10: sig += f"단건{pct(min(AC, 1))} "
    if N > 0 and AD > 0 and AD / N >= 0.3: sig += "SA대체 "
    if AE > 0: sig += f"역분개{AE}건 "
    if z >= 2: sig += "이상치 "
    if I == 0 and F > 0: sig += "신규지출 "
    c1 = min(max(0, N) / 50, 1) * 30; c2 = min(max(0, rate(F, G), rate(F, I)), 1) * 20; c3 = min(max(0, z) / 3, 1) * 25
    qlim = sum((r[6] + r[7]) / 1e6 for r in BP if f"{r[1]}|{r[3]}" == k and r[5] == math.ceil(M / 3))
    qact = tot(k, QS, P); qfc = qact + (mon[P - 2] + mon[P - 1] + mon[P]) / 3 * max(0, QS + 2 - P)
    qrate = qfc / qlim if qlim else 0
    qst = "무예산" if qlim == 0 and qact > 0 else "-" if qlim == 0 else "한도 초과" if qact > qlim else "초과 예상" if qrate > 1 else "주의" if qrate > 0.9 else "여유"
    c4 = 25 if (W in ("부족 예상", "무예산") or qst in ("한도 초과", "초과 예상", "무예산")) else 15 if (W == "주의" or qst == "주의") else 0
    sc = round((c1 + c2 + c3 + c4) * (0.3 if F < 10 and N < 10 else 1), 1)
    K, L = tot(k, QS, P), tot(k, QS - 3, P - 3); AX = tot(k, (Y - 1) * 12 + 1, P - 12)
    VA[k] = dict(d=d, gl=gl, F=F, G=G, I=I, H=rate(F, G), J=rate(F, I), P=avail, Q=Q, U=U, V=V, W=W, X=X, Y=Ym, Z=sig.strip(), AA=sc, mon=mon,
                 K=K, L=L, M=rate(K, L), AX=AX, AY=rate(Q, AX),
                 AJ=qlim, AK=qact, AM=qfc, AN=qrate, AO=qst,
                 row=7 + KEYS.index((d, gl)))
rank = sorted(VA, key=lambda k: VA[k]["AA"] + VA[k]["row"] / 1e6, reverse=True)[:10]

# ---------------- 소진 추세 (선형): 점 (0,0)…(M,누계)에 최소제곱 직선 → 연간가용에 닿는 시점
def xround(x): return math.floor(abs(x) + 0.5) * (1 if x >= 0 else -1)   # 엑셀 ROUND(x,0)
def dstr(d0, t): return (d0 + dt.timedelta(days=xround(t * 365 / 12) - 1)).isoformat()
def burn(cum, avail, ytd, qlim, qact):
    n = M + 1; St = M * (M + 1) / 2; Stt = M * (M + 1) * (2 * M + 1) / 6
    SC = sum(cum); StC = sum((i + 1) * c for i, c in enumerate(cum))
    b = (n * StC - St * SC) / (n * Stt - St ** 2); a = (SC - b * St) / n
    t = "" if (avail <= 0 or ytd >= avail or b <= 0) else max(M, (avail - a) / b)
    st = ("무예산" if avail <= 0 and ytd > 0 else "-" if avail <= 0 else "이미 소진" if ytd >= avail else "소진 안 됨" if b <= 0 else "연내 고갈" if t <= 12 else "연내 여유")
    date = ("소진됨" if st == "이미 소진" else "") if t == "" else dstr(dt.date(Y, 1, 1), t)
    e = P - QS + 1
    if qlim <= 0: qd = ""
    elif qact >= qlim: qd = "한도 소진"
    elif qact <= 0: qd = ""
    elif qlim / (qact / e) > 3: qd = "분기 내 여유"
    else: qd = dstr(dt.date(Y, QS - Y * 12, 1), max(e, qlim / (qact / e)))
    return dict(slope=b, icpt=a, yearEnd=a + b * 12, tstar=t, date=date, status=st, qdate=qd)
for k, v in VA.items():
    v["cum"] = [sum(v["mon"][p] for p in range(Y * 12 + 1, Y * 12 + j + 1)) if j <= M else 0 for j in range(1, 13)]
    v.update(burn(v["cum"], v["P"], v["Q"], v["AJ"], v["AK"]))
    v["AW"] = v["row"] / 1e6 + (0 if v["status"] == "이미 소진" else v["tstar"] if v["status"] == "연내 고갈" else 99)
def burnrow(rank_, dept, account, budget, ytd, b, cum):
    o = dict(rank=rank_, dept=dept, account=account, budget=budget, ytd=ytd, slope=b["slope"], icpt=b["icpt"], yearEnd=b["yearEnd"],
             tstar=b["tstar"], date=b["date"], status=b["status"], qdate=b["qdate"])
    o.update({f"c{j + 1}": c for j, c in enumerate(cum)}); return o
tcum = [sum(v["cum"][j] for v in VA.values()) for j in range(12)]
tav, tytd = sum(v["P"] for v in VA.values()), sum(v["Q"] for v in VA.values())
burnrows = [burnrow("전사", "전사", "사업비 합계", tav, tytd, burn(tcum, tav, tytd, sum(v["AJ"] for v in VA.values()), sum(v["AK"] for v in VA.values())), tcum)]
for i, k in enumerate([k for k in sorted(VA, key=lambda k: VA[k]["AW"]) if VA[k]["AW"] < 99][:15], 1):
    v = VA[k]; burnrows.append(burnrow(i, DEPTS[v["d"]], GLN[v["gl"]], v["P"], v["Q"], v, v["cum"]))

# ---------------- 원인분석 (비교 기준: 전월 · 전분기 · 전년동기누적 · 전년동월 — 원인분석 C10~C13과 같은 기간)
PERIODS = {"전월": (P, P, P - 1, P - 1), "전분기": (QS, P, QS - 3, P - 3), "전년동기누적": (Y * 12 + 1, P, (Y - 1) * 12 + 1, P - 12), "전년동월": (P, P, P - 12, P - 12)}
def cause(k, basis="전월"):
    cs, ce, ps, pe = PERIODS[basis]
    cur = [e for e in E if e["key"] == k and cs <= e["p"] <= ce]; cmp_ = [e for e in E if e["key"] == k and ps <= e["p"] <= pe]
    A16, B16 = sum(e["v"] for e in cur), sum(e["v"] for e in cmp_); C16 = A16 - B16; D16 = rate(A16, B16)
    both = sorted(cur + cmp_, key=lambda e: e["row"]); names = []
    for e in both:
        if e["vname"] not in names: names.append(e["vname"])
    vt = []
    for j, v in enumerate(names[:15], 1):
        a = sum(e["v"] for e in cur if e["vname"] == v); b = sum(e["v"] for e in cmp_ if e["vname"] == v)
        new = "N" if v.startswith("(") else ("Y" if not any(e["key"] == k and e["vname"] == v and cs - 12 <= e["p"] < cs for e in E) else "N")
        vt.append(dict(name=v, cur=a, cmp=b, diff=a - b, nCur=sum(1 for e in cur if e["vname"] == v and e["v"] > 0),
                       nCmp=sum(1 for e in cmp_ if e["vname"] == v and e["v"] > 0), max=max([0] + [e["v"] for e in cur if e["vname"] == v]), new=new, sk=a - b + (31 + j) / 1e6))
    vs = sorted(vt, key=lambda x: x["sk"], reverse=True)
    pats = []
    # 1 일회성 대형
    big = max(cur, key=lambda e: abs(e["v"]) + e["row"] / 1e9) if cur else None
    z45 = sum(e["v"] for e in E if big and e["key"] == k and e["vname"] == big["vname"] and cs - 12 <= e["p"] < cs) / 12
    yes = bool(big) and C16 > 0 and big["v"] > 0 and big["v"] / C16 >= 0.5 and not big["vname"].startswith("(") and big["v"] >= 3 * z45
    sh = min(1, big["v"] / C16) if yes else 0
    pats.append(("일회성 대형 지출", "높음", 3, yes, sh, (f"전표 {big['doc']} ({int(big['date'][5:7])}/{int(big['date'][8:])}, {big['vname']}, 적요 '{big['text']}') {t1(big['v'])}이 증가분의 {pct(sh)} (거래처 평소 월평균 {t1(z45)}). 일회성 지출 가능성") if yes else ""))
    # 2 신규 거래처
    z42 = sum(x["diff"] for x in vt if x["new"] == "Y"); z43 = sum(1 for x in vt if x["new"] == "Y" and x["diff"] > 0)
    z44 = next((x["name"] for x in vs if x["new"] == "Y"), "")
    yes = C16 > 0 and z42 / C16 >= 0.3; sh = min(1, z42 / C16) if yes else 0
    pats.append(("신규 거래처", "중간", 2, yes, sh, f"최근 12개월 거래가 없던 거래처({z44}{' 등 ' + str(z43) + '곳' if z43 > 1 else ''})가 증가분의 {pct(sh)}. 신규 계약·거래 개시 가능성" if yes else ""))
    # 3 단가 인상
    t = vs[0] if vs else None
    yes = bool(t) and C16 > 0 and t["new"] == "N" and t["nCur"] == t["nCmp"] and t["nCmp"] > 0 and t["cmp"] > 0 and 0.05 <= t["diff"] / t["cmp"] <= 0.5
    pats.append(("단가 인상", "중간", 2, yes, min(1, t["diff"] / C16) if yes else 0,
                 f"같은 건수({t['nCur']}건)로 {t['name']} 금액이 {pct(t['diff'] / t['cmp'])} 상승. 단가·계약 조건 변경 가능성 (최근 몇 달 추이 확인)" if yes else ""))
    # 4 시점 이동
    mon = VA[k]["mon"]; a, b, c = mon[P], mon[P - 1], (mon[P - 4] + mon[P - 3] + mon[P - 2]) / 3
    yes = c > 0 and b < 0.5 * c and a > 1.5 * c
    pats.append(("계상 시점 이동", "중간", 2, yes, 0.5 if yes else 0, f"전월 {t1(b)} (직전 3개월 평균 {t1(c)}) → 당월 {t1(a)}. 전월에 계상될 금액이 당월로 넘어왔을 가능성" if yes else ""))
    # 5 정정·역분개
    neg = [e for e in cur if e["v"] < 0]; ns = sum(e["v"] for e in neg)
    yes = len(neg) > 0
    pats.append(("정정 · 역분개", "높음", 3, yes, min(1, abs(ns) / max(abs(C16), 1e-6)) if yes else 0, f"음수(정정·역분개) 전표 {len(neg)}건, 합계 {t1(ns)}. 정정 전후 순액 기준으로 다시 확인" if yes else ""))
    # 6 재분류
    sa = sum(e["v"] for e in cur if e["typ"] == "SA") - sum(e["v"] for e in cmp_ if e["typ"] == "SA")
    satext = next((e["text"] for e in sorted(cur, key=lambda e: e["row"]) if e["typ"] == "SA"), "")
    yes = C16 > 0 and sa / C16 >= 0.3; sh = min(1, sa / C16) if yes else 0
    pats.append(("재분류 (수기 대체)", "높음", 3, yes, sh, f"수기 대체전표(SA) 순증 {t1(sa)}이 증가분의 {pct(sh)}. 다른 부서·계정에서 이관됐을 가능성 (적요: '{satext}')" if yes else ""))
    # 7 적요 키워드
    best = None
    for idx, (nm, ks, desc) in enumerate(KW):
        amt = max([0] + [sum(e["v"] for e in cur if kk in e["text"]) for kk in ks if kk])
        share = amt / C16 if C16 > 0 else None
        if share is not None and (best is None or share > best[0]): best = (share, idx, amt)
    z34 = best[0] if best else 0
    yes = C16 > 0 and z34 > 0 and z34 >= 0.3
    txt = ""
    if yes:
        nm, ks, desc = KW[best[1]]
        kl = ks[0] + ("·" + ks[1] if ks[1] else "") + ("·" + ks[2] if ks[2] else "")
        txt = f"{nm}: 적요에 '{kl}' 등이 있는 전표 {t1(best[2])} (증가분의 {pct(z34)}). {desc}"
    pats.append(("적요 키워드", "중간", 2, yes, min(1, z34) if yes else 0, txt))
    # 8 물량 vs 단가
    n1 = sum(1 for e in cur if e["v"] > 0); n0 = sum(1 for e in cmp_ if e["v"] > 0)
    s1 = sum(e["v"] for e in cur if e["v"] > 0); s0 = sum(e["v"] for e in cmp_ if e["v"] > 0)
    a1 = s1 / n1 if n1 else 0; a0 = s0 / n0 if n0 else 0
    yes = C16 > 0 and n1 > 0 and n0 > 0
    pats.append(("물량 vs 단가", "참고", 1, yes, 0, f"건수 {n0}→{n1} ({spct(n1 / max(n0, 1) - 1)}), 건당 {t1(a0)}→{t1(a1)}. {'건수(물량) 증가' if abs((n1 - n0) * a0) >= abs((a1 - a0) * n1) else '건당 금액(단가) 상승'} 영향이 더 큼" if yes else ""))
    cands = sorted([(sc * 10 + sh + (20 + i) / 1e6, nm, lv, txt) for i, (nm, lv, sc, yes, sh, txt) in enumerate(pats) if yes], reverse=True)
    top = cands[0][3] if cands else "근거 전표 목록을 직접 확인하세요"
    va = VA[k]
    summ = (f"{DEPTS[va['d']]} · {GLN[va['gl']]} {basis} 대비 {('+' if C16 > 0 else '') + t1(C16) if C16 != 0 else '0'} ({spct(D16)}). "
            + ("감소 또는 변동 없음 - 원인 후보는 증가 항목에만 만듭니다. " if C16 <= 0 else top + ". ")
            + f"분기 한도: {va['AO']}" + (f", 연간 {t1(va['X'])} 부족 예상 ({va['Y']} 소진)" if va["X"] > 0 else "") + ".")
    ev = sorted(cur, key=lambda e: abs(e["v"]) + e["row"] / 1e9, reverse=True)[:15]
    return dict(basis=basis, summary=summ, curTotal=A16, cmpTotal=B16, diff=C16,
                candidates=[dict(pattern=c[1], strength=c[2], text=c[3]) for c in cands] or [dict(pattern="뚜렷한 패턴 없음", strength=None, text="근거 전표 목록을 직접 확인하세요")],
                vendors=[dict(name=x["name"], cur=x["cur"], cmp=x["cmp"], diff=x["diff"], share=(x["diff"] / C16 if C16 > 0 else None), isNew=x["new"]) for x in vs],
                evidence=[dict(date=e["date"], doc=e["doc"], type=e["typ"], vendor=e["vname"], text=e["text"], amt=e["v"]) for e in ev])

top10 = []
for i, k in enumerate(rank, 1):
    v = VA[k]
    bw = dict(dept=DEPTS[v["d"]], account=GLN[v["gl"]], ytd=v["Q"], budget=v["P"], slope=v["slope"], icpt=v["icpt"], yearEnd=v["yearEnd"],
              tstar=v["tstar"], date=v["date"], status=v["status"], qdate=v["qdate"]); bw.update({f"c{j + 1}": c for j, c in enumerate(v["cum"])})
    top10.append(dict(rank=i, dept=DEPTS[v["d"]], account=GLN[v["gl"]], cur=v["F"], mom=v["H"], yoy=v["J"], status=v["AO"], qrate=v["AN"],
                      annual=v["W"], signal=v["Z"] or None, score=v["AA"], comment=COMMENTS.get(k, "(미확인)"),
                      qcur=v["K"], qoq=v["M"], ytd=v["Q"], ytdRate=v["AY"], burn=bw,
                      cause=cause(k), causes={b: cause(k, b) for b in ("전분기", "전년동기누적")}))

# ---------------- 차트데이터
bs = {(r[0] * 12 + r[1], r[2]): r[6] / 1e6 for r in D["BS"]}
pl = {(r[0] * 12 + r[1], r[2]): (r[6] / 1e6, r[7] / 1e6) for r in D["PL"]}
avail_q = lambda q: sum((r[6] + r[7]) / 1e6 for r in BP if r[5] == q)
monthly = [dict(month=f"{m}월", cur=(sum(e["v"] for e in E if e["p"] == Y * 12 + m) if m <= M else 0),
                prev=sum(e["v"] for e in E if e["p"] == (Y - 1) * 12 + m), budget=avail_q(math.ceil(m / 3)) / 3) for m in range(1, 13)]
depts = []
for dd, nm in DEPTS.items():
    a = sum(v["AJ"] for v in VA.values() if v["d"] == dd); q = sum(v["AK"] for v in VA.values() if v["d"] == dd); u = sum(v["AM"] for v in VA.values() if v["d"] == dd)
    r_ = u / a if a else 0
    depts.append(dict(code=dd, name=nm, avail=a, ytd=q, forecast=u, rate=r_, status="-" if a == 0 else "한도 초과" if q > a else "초과 예상" if r_ > 1 else "주의" if r_ > 0.9 else "여유"))
status = [dict(status=s, count=sum(1 for v in VA.values() if v["AO"] == s)) for s in ["여유", "주의", "초과 예상", "한도 초과", "무예산"]]
accounts = []   # 전표 기준 (변동분석 행 합계) — 엑셀 차트데이터 계정별과 같음
for gl, nm in GLN.items():
    vs = [v for v in VA.values() if v["gl"] == gl]; sm = lambda f: sum(v[f] for v in vs)
    accounts.append(dict(gl=gl, name=nm, cur=sm("F"), prevYear=sm("I"), diff=sm("F") - sm("I"), prevMonth=sm("G"), qCur=sm("K"), qPrev=sm("L"), ytd=sm("Q"), ytdPrev=sm("AX")))
bsrows = [dict(p=p, period=plab(p), assets=bs[(p, "A000")], liabilities=bs[(p, "L000")], equity=bs[(p, "E000")], csm=bs[(p, "L130")], invested=bs[(p, "A100")]) for p in range(P - 11, P + 1)]
plrows = [dict(month=f"{m}월", niCur=pl[(Y * 12 + m, "P600")][0] if m <= M else 0, niPrev=pl[((Y - 1) * 12 + m, "P600")][0],
               insCur=pl[(Y * 12 + m, "P100")][0] if m <= M else 0, invCur=pl[(Y * 12 + m, "P200")][0] if m <= M else 0,
               expCur=pl[(Y * 12 + m, "P300")][0] if m <= M else 0, insPrev=pl[((Y - 1) * 12 + m, "P100")][0]) for m in range(1, 13)]
ytd = sum(e["v"] for e in E if Y * 12 + 1 <= e["p"] <= P); pytd = sum(e["v"] for e in E if (Y - 1) * 12 + 1 <= e["p"] <= P - 12)
mc, mp = sum(e["v"] for e in E if e["p"] == P), sum(e["v"] for e in E if e["p"] == P - 1)
kv = lambda a, b: (b and a / b - 1) or 0
kpi = [dict(label="총자산", value=bs[(P, "A000")], compare=bs[(P - 12, "A000")], rate=kv(bs[(P, "A000")], bs[(P - 12, "A000")]), basis="전년동월말"),
       dict(label="자본총계", value=bs[(P, "E000")], compare=bs[(P - 12, "E000")], rate=kv(bs[(P, "E000")], bs[(P - 12, "E000")]), basis="전년동월말"),
       dict(label="보험계약마진(CSM)", value=bs[(P, "L130")], compare=bs[(P - 12, "L130")], rate=kv(bs[(P, "L130")], bs[(P - 12, "L130")]), basis="전년동월말"),
       dict(label="당기순이익 누계", value=pl[(P, "P600")][1], compare=pl[(P - 12, "P600")][1], rate=kv(pl[(P, "P600")][1], pl[(P - 12, "P600")][1]), basis="전년 누계"),
       dict(label="사업비 누계", value=ytd, compare=pytd, rate=kv(ytd, pytd), basis="전년 누계"),
       dict(label="사업비 당월", value=mc, compare=mp, rate=kv(mc, mp), basis="전월"),
       dict(label="사업비 당분기", value=sum(v["K"] for v in VA.values()), compare=sum(v["L"] for v in VA.values()),
            rate=kv(sum(v["K"] for v in VA.values()), sum(v["L"] for v in VA.values())), basis="전분기"),
       dict(label="분기 한도 초과(예상)", value=status[2]["count"] + status[3]["count"], compare=status[1]["count"], rate=None, basis="주의 건수")]
# ---------------- IFRS17 (GMM) 사업비 영향
DIRECT = {"63501000": "신계약비", "63502000": "수금비", "62701000": "비귀속"}
NATURE_OTHER = {"62601000", "62801000", "62802000"}
DEPT_ALLOC = {"D100": (0, .10, 0, .90), "D200": (0, .20, .10, .70), "D300": (.05, .15, 0, .80), "D400": (.50, .10, 0, .40),
              "D500": (.15, 0, 0, .85), "D600": (.75, .10, 0, .15), "D700": (.10, .40, .10, .40), "D800": (0, .60, .30, .10),
              "D900": (.90, .05, .05, 0), "D950": (0, .30, 0, .70)}
INV_DEPTS = ["D950"]; INV_RATE = {2025: 0.65, 2026: 0.70}
def alloc_of(e):   # [비귀속, 신계약, 유지, 수금, 재산관리비]
    if e["gl"] in DIRECT:
        return [1.0 if DIRECT[e["gl"]] == c else 0.0 for c in ("비귀속", "신계약비", "유지비", "수금비")] + [0.0]
    dd = CC2D[e["cc"]]
    inv = INV_RATE.get(e["y"], 0) if (e["gl"] not in NATURE_OTHER and dd in INV_DEPTS) else 0.0
    sn, mt, cl, na = DEPT_ALLOC.get(dd, (0, 0, 0, 1))
    return [(1 - inv) * na, (1 - inv) * sn, (1 - inv) * mt, (1 - inv) * cl, inv]
PFS = [("PF01", "보장성", "GMM"), ("PF02", "저축성", "GMM"), ("PF03", "단체(단기)", "PAA"), ("PF04", "변액", "VFA")]
PF_DRV = {"PF01": (.55, .50, .50), "PF02": (.20, .30, .30), "PF03": (.10, .08, .08), "PF04": (.15, .12, .12)}
for e in E:
    e["al"] = [e["v"] * x for x in alloc_of(e)]
def catsum(j, p0, p1): return sum(e["al"][j] for e in E if p0 <= e["p"] <= p1)
CATS = ["비귀속", "신계약비", "유지비", "수금비", "재산관리비"]
alloc = [dict(cat=c, cur=catsum(j, P, P), prev=catsum(j, P - 1, P - 1), ytd=catsum(j, Y * 12 + 1, P)) for j, c in enumerate(CATS)]
tot_cur = sum(a["cur"] for a in alloc)
for a in alloc: a["share"] = a["cur"] / tot_cur if tot_cur else 0
EXPD = {(r[0] * 12 + r[1], r[2]): [x / 1e6 for x in r[4:7]] for r in D["EXP"]}
def target(model, cat):
    if cat == "신계약비" and model in ("GMM", "VFA"): return f"CSM 조정 ({model})"
    if cat == "신계약비": return "당기손익"
    return "당기손익(보험서비스결과)"
pfrows = []
for pf, pfn, mdl in PFS:
    for j, cat in enumerate(CATS[1:4]):
        act = alloc[j + 1]["cur"] * PF_DRV[pf][j]; exp_ = EXPD[(P, pf)][j]
        pfrows.append(dict(pf=pf, name=pfn, model=mdl, cat=cat, actual=act, expected=exp_, diff=exp_ - act, target=target(mdl, cat)))
tcsm = sum(r["diff"] for r in pfrows if r["target"] == "CSM 조정 (GMM)")
tvfa = sum(r["diff"] for r in pfrows if r["target"] == "CSM 조정 (VFA)")
tpl = sum(r["diff"] for r in pfrows if r["target"].startswith("당기손익"))
ROLLP = [r for r in D["ROLL"] if r[0] * 12 + r[1] == P]
recon = []
for pf, pfn, mdl in PFS:
    act = sum(r[9] for r in ROLLP if r[2] == pf and r[5] == "F31") / 1e6
    tr = 0 if mdl == "PAA" else sum(r["diff"] for r in pfrows if r["pf"] == pf and r["cat"] == "신계약비")
    recon.append(dict(pf=pf, model=mdl, actuary=act, trace=tr, diff=act - tr, status="해당없음 (CSM 없음)" if mdl == "PAA" else ("일치" if abs(act - tr) < 0.5 else "차이 확인")))
FACN = [("F00", "기초잔액"), ("F10", "신계약 효과"), ("F21", "가정변경-사업비"), ("F22", "가정변경-기타"), ("F31", "경험조정-신계약비"),
        ("F32", "경험조정-보험료 등"), ("F40", "보험금융비용(이자부리)"), ("F51", "CSM 상각"), ("F52", "RA 해제"), ("F60", "현금흐름 및 기타"), ("F99", "기말잔액")]
def mkbridge(mdl):
    return [dict(code=c, name=n, bel=sum(r[7] for r in ROLLP if r[5] == c and r[4] == mdl) / 1e6, ra=sum(r[8] for r in ROLLP if r[5] == c and r[4] == mdl) / 1e6,
                 csm=sum(r[9] for r in ROLLP if r[5] == c and r[4] == mdl) / 1e6, expense=c in ("F21", "F31")) for c, n in FACN]
bridge, bridgeV = mkbridge("GMM"), mkbridge("VFA")
bmap = {b["code"]: b for b in bridge}; vmap = {b["code"]: b for b in bridgeV}
csm_chg = bmap["F99"]["csm"] - bmap["F00"]["csm"]; csm_exp = bmap["F21"]["csm"] + bmap["F31"]["csm"]
vcsm_chg = vmap["F99"]["csm"] - vmap["F00"]["csm"]; vcsm_exp = vmap["F21"]["csm"] + vmap["F31"]["csm"]
drv = []
for i, (d_, gl_) in enumerate(KEYS):
    k_ = f"{d_}|{gl_}"
    c_ = sum(e["al"][1] for e in E if e["key"] == k_ and e["p"] == P); p_ = sum(e["al"][1] for e in E if e["key"] == k_ and e["p"] == P - 1)
    drv.append(dict(dept=DEPTS[d_], account=GLN[gl_], cur=c_, prev=p_, diff=c_ - p_, sk=c_ - p_ + (7 + i) / 1e6))
drivers = sorted(drv, key=lambda x: x["sk"], reverse=True)[:5]
evn = sorted([e for e in E if e["p"] == P and e["al"][1] > 0], key=lambda e: e["al"][1] + e["row"] / 1e9, reverse=True)[:10]
evidence17 = [dict(date=e["date"], doc=e["doc"], cc=e["cc"], gl=e["gl"], vendor=e["vname"], text=e["text"], amt=e["v"], nb=e["al"][1]) for e in evn]
def sg(x, d=1):
    v = round(x, d)
    if v == 0: return "0"
    return ("+" if v > 0 else "-") + f"{abs(v):,.{d}f}"
nbt = sum(r["diff"] for r in pfrows if r["cat"] == "신계약비")
summary17 = (f"GMM CSM 당월 변동 {sg(csm_chg, 0)} 중 사업비 관련 {sg(csm_exp)} (신계약비 경험조정 {sg(bmap['F31']['csm'])}, 사업비 가정변경 {sg(bmap['F21']['csm'])}). "
             f"신계약비 실제가 예정보다 {abs(nbt):,.1f}{' 많았고' if nbt < 0 else ' 적었고'}"
             + (f", 증가 기여 1위는 {drivers[0]['dept']} {drivers[0]['account']} ({sg(drivers[0]['diff'])})" if drivers else "")
             + (f", 대표 전표는 {evidence17[0]['vendor']} '{evidence17[0]['text']}' {evidence17[0]['nb']:,.1f}" if evidence17 else "")
             + f". 당기손익 반영 경험차이 {sg(tpl)}. VFA(변액) CSM 당월 변동 {sg(vcsm_chg, 0)} 중 사업비 관련 {sg(vcsm_exp)}.")
ifrs17 = dict(summary=summary17, alloc=alloc, pf=pfrows, totals=dict(csm=tcsm, csmVfa=tvfa, pl=tpl), recon=recon, bridge=bridge, bridgeVfa=bridgeV,
              csmChange=csm_chg, csmExpense=csm_exp, csmChangeVfa=vcsm_chg, csmExpenseVfa=vcsm_exp, drivers=[{k: v for k, v in d.items() if k != "sk"} for d in drivers], evidence=evidence17)

data = dict(meta=dict(company=os.environ.get("COMPANY", "예시보험㈜"), period=plab(P), unit="백만원", check="전체 정상 (참고 1건)", recon="전체 일치", generated="2026-10-14 (예시 데이터)"),
            kpi=kpi, monthly=monthly, depts=depts, status=status, accounts=accounts, bs=bsrows, pl=plrows, burn=burnrows, top10=top10, ifrs17=ifrs17)
def rnd(o):
    if isinstance(o, float): return round(o, 4)
    if isinstance(o, dict): return {k: rnd(v) for k, v in o.items()}
    if isinstance(o, list): return [rnd(v) for v in o]
    return o
open(OUT, "w", encoding="utf-8").write("window.TRACE_DATA = " + json.dumps(rnd(data), ensure_ascii=False, indent=1) + ";\n")
print("wrote", OUT)
for t in top10: print(t["rank"], t["dept"], t["account"], t["status"], t["signal"], "|", [c["pattern"] for c in t["cause"]["candidates"]])
for b in burnrows: print("소진", b["rank"], b["dept"], b["account"], round(b["budget"], 1), round(b["ytd"], 1), round(b["slope"], 2), b["tstar"] if b["tstar"] == "" else round(b["tstar"], 2), b["date"], b["status"], b["qdate"])
