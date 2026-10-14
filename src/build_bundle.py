"""TRACE 전체 묶음(ZIP) 생성: 내부망 반입용 파일을 한 번에 내려받도록 묶는다.

사용: python src/build_bundle.py   (다른 생성 스크립트 다음에 실행) → TRACE_v0.5_전체.zip
"""
import os, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = ["TRACE_v0.5.xlsm", "TRACE_v0.5.xlsx", "modTrace.bas", "TRACE_대시보드.html", "data.js", "TRACE_흐름.html", "TRACE_IFRS17_연결.html",
         "TRACE_사용설명서.html", "TRACE_내려받기.html", "trace.css", "samples/TRACE_입력양식_v0.5.xlsx", "samples/샘플_감독원양식_2026-09.xlsx"]
OUT = os.path.join(ROOT, "TRACE_v0.5_전체.zip")

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for f in FILES:
        info = zipfile.ZipInfo.from_file(os.path.join(ROOT, f), "TRACE_v0.5/" + f)
        info.date_time = (2026, 10, 1, 0, 0, 0)        # 같은 내용이면 같은 ZIP이 나오도록 시각 고정
        info.flag_bits |= 0x800                         # 파일 이름 UTF-8 표시 (Windows 탐색기에서 한글 깨짐 방지)
        info.compress_type = zipfile.ZIP_DEFLATED
        with open(os.path.join(ROOT, f), "rb") as fh: z.writestr(info, fh.read())
print("saved", OUT, os.path.getsize(OUT), "bytes")

# 아티팩트(claude.ai)는 .xlsx·.zip을 파일로 게시할 수 없어서, 내려받기 페이지가 읽을 base64 스크립트로도 만든다 (dl/, 저장소에는 올리지 않음)
# 순서 = dl/f0~f4.js (내려받기 페이지의 data-src와 맞춤). 아티팩트 저장 기능은 .xlsm을 받지 않아 ZIP에 담는다
with zipfile.ZipFile(os.path.join(ROOT, "dl_xlsm.zip"), "w", zipfile.ZIP_DEFLATED) as z:
    info = zipfile.ZipInfo("TRACE_v0.5.xlsm", (2026, 10, 1, 0, 0, 0)); info.compress_type = zipfile.ZIP_DEFLATED
    z.writestr(info, open(os.path.join(ROOT, "TRACE_v0.5.xlsm"), "rb").read())
DL = {"TRACE_v0.5.xlsx": "TRACE_v0.5.xlsx", "TRACE_입력양식_v0.5.xlsx": "samples/TRACE_입력양식_v0.5.xlsx",
      "샘플_감독원양식_2026-09.xlsx": "samples/샘플_감독원양식_2026-09.xlsx", "TRACE_v0.5_전체.zip": "TRACE_v0.5_전체.zip",
      "TRACE_v0.5.xlsm.zip": "dl_xlsm.zip"}
import base64, json
os.makedirs(os.path.join(ROOT, "dl"), exist_ok=True)
for i, (name, src) in enumerate(DL.items()):
    with open(os.path.join(ROOT, src), "rb") as fh: b64 = base64.b64encode(fh.read()).decode()
    with open(os.path.join(ROOT, "dl", f"f{i}.js"), "w", encoding="utf-8") as fh:
        fh.write("(window.TRACE_FILES=window.TRACE_FILES||{})[%s]=\"%s\";\n" % (json.dumps(name, ensure_ascii=False), b64))
os.remove(os.path.join(ROOT, "dl_xlsm.zip"))
print("saved dl/f0~f%d.js" % (len(DL) - 1))
