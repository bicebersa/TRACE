"""TRACE 수식 검증: 작은 범위로 엑셀을 만들고 formulas 라이브러리로 전 수식을 계산해 파이썬 대조값(build_datajs.py)과 비교한다.

사용:
  pip install openpyxl formulas
  python tools/verify.py                         # 원인분석 1순위 기준
  python tools/verify.py --sel D200,62901000,전월  # 원인분석을 특정 부서·계정으로 지정해 패턴 문장까지 확인

엑셀·LibreOffice 없이 검증하는 용도다. 메모리 8GB면 한 번에 하나만 실행한다 (10분 안팎).
formulas는 TEXT(x,"+0%")의 + 부호와 TEXT(0,"#,##0.0")를 잘못 표시하므로 그 차이는 무시해도 된다.
"""
import argparse, contextlib, io, json, os, re, runpy, subprocess, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")
OUT = os.path.join(ROOT, "tools", "_out")

EVAL = r'''
import formulas, json, os, sys
xl = formulas.ExcelModel().loads(sys.argv[1]).finish()
sol = xl.calculate()
out = {}
for k, v in sol.items():
    try: val = v.value[0][0]
    except Exception: val = v
    out[k.upper()] = str(val)
with open(sys.argv[2], "w", encoding="utf-8") as fh: json.dump(out, fh, ensure_ascii=False)
sys.stdout.flush(); os._exit(0)
'''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sel", default="", help="원인분석 직접 지정: 부서코드,G/L,비교기준")
    a = ap.parse_args()
    os.makedirs(OUT, exist_ok=True)
    env = dict(os.environ, JE_END="6000", SEL=a.sel)
    print("1/3 엑셀 생성 (전표 범위 6000행)")
    subprocess.run([sys.executable, os.path.join(SRC, "build_trace.py"), OUT], check=True, env=env)
    book = [f for f in os.listdir(OUT) if f.startswith("TRACE_v") and f.endswith(".xlsx")][0]
    sol = os.path.join(OUT, "sol.json")
    print("2/3 수식 계산 (수 분 걸림)")
    with open(os.path.join(OUT, "_eval.py"), "w", encoding="utf-8") as fh: fh.write(EVAL)
    rr = subprocess.run([sys.executable, os.path.join(OUT, "_eval.py"), os.path.join(OUT, book), sol], stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace")
    if rr.returncode != 0:   # 계산 단계 오류는 원인 몇 줄만 보여 주고 멈춤 (예: formulas 미설치)
        print("수식 계산 실패. 마지막 오류:\n" + "\n".join(rr.stderr.strip().splitlines()[-6:]))
        if "No module named" in rr.stderr: print("→ 'py -m pip install openpyxl formulas' 를 실행한 뒤 다시 해 보세요.")
        sys.exit(1)
    print("3/3 대조")
    o = json.load(open(sol, encoding="utf-8"))
    tag = "[" + book.upper() + "]"
    def g(sh, c):
        v = o.get(f"'{tag}{sh}'!{c}")
        try: return float(v)
        except (TypeError, ValueError): return v
    errs = [k for k, v in o.items() if str(v).startswith("#")]
    sys.argv = ["build_datajs.py", os.devnull]
    with contextlib.redirect_stdout(io.StringIO()):
        G = runpy.run_path(os.path.join(SRC, "build_datajs.py"))
    VA, data, cause = G["VA"], G["data"], G["cause"]
    bad = []
    for r in range(7, 120):
        k = g("변동분석", f"E{r}")
        if k not in VA: continue
        x = VA[k]
        for col, f in (("F", "F"), ("G", "G"), ("I", "I"), ("P", "P"), ("Q", "Q"), ("U", "U"), ("AA", "AA"), ("AJ", "AJ"), ("AM", "AM"), ("K", "K"), ("L", "L"), ("M", "M"), ("AX", "AX"), ("AY", "AY"), ("AZ", "AZ")):
            v = g("변동분석", f"{col}{r}")
            if not isinstance(v, float) or abs(v - x[f]) > 0.06: bad.append(f"변동분석 {k} {col}: {v} vs {x[f]}")
        for col, f in (("W", "W"), ("AO", "AO")):
            if g("변동분석", f"{col}{r}") != x[f]: bad.append(f"변동분석 {k} {col}: {g('변동분석', f'{col}{r}')} vs {x[f]}")
        for col, f in (("AP", "slope"), ("AQ", "icpt"), ("AS", "tstar")):   # 소진 추세 (선형)
            v = g("변동분석", f"{col}{r}")
            if x[f] == "":
                if v not in ("", None): bad.append(f"변동분석 {k} {col}: {v} vs ''")
            elif not isinstance(v, float) or abs(v - x[f]) > 0.06: bad.append(f"변동분석 {k} {col}: {v} vs {x[f]}")
        for col, f in (("AT", "date"), ("AU", "status"), ("AV", "qdate")):
            v = g("변동분석", f"{col}{r}")
            if (v or "") != x[f]: bad.append(f"변동분석 {k} {col}: {v} vs {x[f]}")
    CDL = json.load(open(os.path.join(SRC, "tool_data.json")))["CD"]
    for i, t in enumerate(data["top10"]):   # 주의대상 누계 비교·고갈 추세 열
        r = 6 + i
        for col, v in (("O", t["qcur"]), ("P", t["qoq"]), ("Q", t["ytd"]), ("R", t["ytdRate"]), ("S", t["burn"]["budget"]), ("T", t["burn"]["slope"]), ("AL", t["burn"]["c12"]), ("AA", t["burn"]["c1"])):
            x = g("주의대상", f"{col}{r}")
            if not isinstance(x, float) or abs(x - v) > 0.06: bad.append(f"주의대상 {i + 1} {col}: {x} vs {v}")
        for col, v in (("X", t["burn"]["date"]), ("Y", t["burn"]["status"]), ("Z", t["burn"]["qdate"])):
            if (g("주의대상", f"{col}{r}") or "") != v: bad.append(f"주의대상 {i + 1} {col}: {g('주의대상', f'{col}{r}')} vs {v}")
    for i, a in enumerate(data["accounts"]):   # 차트데이터 계정별 (전표 기준)
        r = CDL["CDD"] + 1 + i
        for col, f in (("C", "cur"), ("D", "prevYear"), ("F", "prevMonth"), ("G", "qCur"), ("H", "qPrev"), ("I", "ytd"), ("J", "ytdPrev")):
            x = g("차트데이터", f"{col}{r}")
            if not isinstance(x, float) or abs(x - a[f]) > 0.06: bad.append(f"차트데이터 계정 {a['gl']} {col}: {x} vs {a[f]}")
    for i, t in enumerate(data["top10"]):   # 주의대상 예산 조정·조치 열
        r = 6 + i
        v = g("주의대상", f"AQ{r}")
        if not isinstance(v, (int, float)) or abs(v - t["adjQ"]) > 0.06: bad.append(f"주의대상 {i + 1} AQ: {v} vs {t['adjQ']}")
        for col, f in (("AR", "adjType"), ("AT", "lastAction")):
            if (g("주의대상", f"{col}{r}") or "") != t[f]: bad.append(f"주의대상 {i + 1} {col}: {g('주의대상', f'{col}{r}')} vs {t[f]}")
        v = g("주의대상", f"AS{r}")
        if not isinstance(v, (int, float)) or int(v) != t["actCount"]: bad.append(f"주의대상 {i + 1} AS: {v} vs {t['actCount']}")
    for i, b in enumerate(data["budgetHistory"]):   # 예산변경이력 보조수식
        r = 5 + i
        for col, f in (("J", "key"), ("M", "from"), ("N", "to"), ("Q", "type"), ("R", "move")):
            if (g("예산변경이력", f"{col}{r}") or "") != (b[f] or ""): bad.append(f"예산변경이력 {r} {col}: {g('예산변경이력', f'{col}{r}')} vs {b[f]}")
        for col, f in (("O", "fromAmt"), ("P", "amt")):
            v = g("예산변경이력", f"{col}{r}")
            if not isinstance(v, (int, float)) or abs(v - b[f]) > 0.06: bad.append(f"예산변경이력 {r} {col}: {v} vs {b[f]}")
    BPD = json.load(open(os.path.join(SRC, "tool_data.json")))["BP"]
    sys.path.insert(0, SRC)
    for i, row in enumerate(BPD):   # 예산판 조정 추정·조치 기록
        if row[7] == 0: continue
        r = 5 + i
        exp_t = G["adjtype"](row); rec = any(c[2] == row[1] and str(c[3]) == str(row[3]) and c[0] == row[0] and c[10] == row[5] and c[8] for c in G["CM"])
        if g("예산판_원본", f"S{r}") != exp_t: bad.append(f"예산판 {r} S: {g('예산판_원본', f'S{r}')} vs {exp_t}")
        if g("예산판_원본", f"T{r}") != ("기록됨" if rec else "기록 없음"): bad.append(f"예산판 {r} T: {g('예산판_원본', f'T{r}')} vs {rec}")
    for i, x_ in enumerate(data["expcat"]):   # 차트데이터 사업비 구성 (재무 요약)
        r = CDL["CDI"] + 1 + i
        for col, f in [(chr(66 + m), f"m{m + 1}") for m in range(12)] + [("N", "ytd"), ("P", "py"), ("Q", "diff")]:
            v = g("차트데이터", f"{col}{r}")
            if not isinstance(v, (int, float)) or abs(v - x_[f]) > 0.06: bad.append(f"차트데이터 사업비구성 {x_['cat']} {col}: {v} vs {x_[f]}")
    tot_ = data["expcat"][-1]["ytd"]; kp_ = next(k for k in data["kpi"] if k["label"] == "사업비 누계")["value"]
    if abs(tot_ - kp_) > 0.06: bad.append(f"사업비 구성 합계 {tot_} vs 사업비 누계 {kp_}")
    for i, kp in enumerate(data["kpi"]):
        x = g("차트데이터", f"B{CDL['CDG'] + 1 + i}")
        if not isinstance(x, float) or abs(x - kp["value"]) > 0.6: bad.append(f"KPI {kp['label']}: {x} vs {kp['value']}")
    for i, b in enumerate(G["burnrows"]):
        r = CDL["CDH"] + 1 + i
        for col, f in (("A", "rank"), ("B", "dept"), ("J", "date"), ("K", "status"), ("L", "qdate")):
            v = g("차트데이터", f"{col}{r}"); v = int(v) if isinstance(v, float) and f == "rank" else v
            if v != b[f]: bad.append(f"차트데이터 소진 {i} {col}: {v} vs {b[f]}")
        for col, f in (("D", "budget"), ("E", "ytd"), ("F", "slope"), ("G", "icpt"), ("X", "c12"), ("M", "c1")):
            v = g("차트데이터", f"{col}{r}")
            if not isinstance(v, float) or abs(v - b[f]) > 0.06: bad.append(f"차트데이터 소진 {i} {col}: {v} vs {b[f]}")
    top_xl = [g("주의대상", f"L{r}") for r in range(6, 16)]
    top_py = [f"{VA[k]['d']}|{VA[k]['gl']}" for k in G["rank"]]
    if top_xl != top_py: bad.append(f"TOP10 {top_xl} vs {top_py}")
    key = g("원인분석", "C9"); basis = g("원인분석", "C6"); py = cause(key, basis)
    xl_c = [g("원인분석", f"B{r}") for r in range(21, 29) if g("원인분석", f"B{r}")]
    if xl_c != [c["pattern"] for c in py["candidates"]]: bad.append(f"원인 패턴 {xl_c} vs {[c['pattern'] for c in py['candidates']]}")
    X = data["ifrs17"]; L = json.load(open(os.path.join(SRC, "tool_data.json")))["IF"]
    for i, r in enumerate(X["pf"]):
        v = g("IFRS17_사업비", f"G{L['B0'] + 1 + i}")
        if not isinstance(v, float) or abs(v - r["diff"]) > 0.06: bad.append(f"IFRS17 경험차이 {r['pf']} {r['cat']}: {v} vs {r['diff']}")
        for col, f in (("I", "csm"), ("J", "pl")):   # CSM 반영 · 당기손익 반영
            v = g("IFRS17_사업비", f"{col}{L['B0'] + 1 + i}"); v = 0.0 if v in ("", None) else v
            if not isinstance(v, float) or abs(v - r[f]) > 0.06: bad.append(f"IFRS17 {f} {r['pf']} {r['cat']}: {v} vs {r[f]}")
        hx = g("IFRS17_사업비", "H%d" % (L["B0"] + 1 + i))
        if hx != r["target"]: bad.append("IFRS17 반영처 %s %s: %s vs %s" % (r["pf"], r["cat"], hx, r["target"]))
    print(f"\n수식 오류: {len(errs)}건 {errs[:5]}")
    print(f"자가진단: {g('자가진단', 'B4')}")
    print(f"원인분석 선택: {key} · 원인 후보: {xl_c}")
    print(f"대조 불일치: {len(bad)}건")
    for b in bad[:30]: print("  ", b)
    sys.exit(1 if errs or bad else 0)


if __name__ == "__main__":
    main()
