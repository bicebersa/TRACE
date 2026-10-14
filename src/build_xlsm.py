"""TRACE_v0.5.xlsx + modTrace.bas → 매크로 내장 TRACE_v0.5.xlsm (열면 설정 시트에 버튼이 생김).

엑셀 없이 MS-OVBA 규격대로 vbaProject.bin을 직접 만든다.
- 소스 코드만 넣고 컴파일 캐시(p-code)는 넣지 않는다 (_VBA_PROJECT Version 0xFFFF → 엑셀이 열 때 소스에서 컴파일).
- 모듈: ThisWorkbook(Workbook_Open → EnsureButtons), 시트마다 문서 모듈(Sheet1…), modTrace(표준 모듈).
- 한글 소스는 CP949(코드 페이지 949)로 저장한다.

사용: python src/build_xlsm.py   (build_trace.py · build_vba.py 다음에 실행)
"""
import io, os, random, re, struct, sys, uuid, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_XLSX = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "TRACE_v0.5.xlsx")
SRC_BAS = sys.argv[2] if len(sys.argv) > 2 else os.path.join(ROOT, "modTrace.bas")
OUT = sys.argv[3] if len(sys.argv) > 3 else os.path.join(ROOT, "TRACE_v0.5.xlsm")
CP = "cp949"; CODEPAGE = 949


# ---------------------------------------------------------------- MS-OVBA 압축 (2.4.1)
def _copy_token_help(diff):
    bits = max((diff - 1).bit_length(), 4)
    return bits, (0xFFFF >> bits) + 3          # (offset 비트 수, 최대 길이)

def _compress_chunk(data):
    out = bytearray(); i, n = 0, len(data); index = {}
    def add(pos):
        if pos + 3 <= n: index.setdefault(data[pos:pos + 3], []).append(pos)
    while i < n:
        flag_pos = len(out); out.append(0); flag = 0
        for bit in range(8):
            if i >= n: break
            bits, maxlen = _copy_token_help(i) if i > 0 else (4, 0)
            best_len, best_off = 0, 0
            if i > 0:
                for cand in reversed(index.get(data[i:i + 3], [])[-64:]):
                    L = 0; lim = min(maxlen, n - i)
                    while L < lim and data[cand + L] == data[i + L]: L += 1
                    if L > best_len: best_len, best_off = L, i - cand
                    if L == lim: break
            if best_len >= 3:
                tok = ((best_off - 1) << (16 - bits)) | (best_len - 3)
                out += struct.pack("<H", tok); flag |= 1 << bit
                for k in range(best_len): add(i + k)
                i += best_len
            else:
                out.append(data[i]); add(i); i += 1
        out[flag_pos] = flag
    return bytes(out)

def ovba_compress(data):
    res = bytearray(b"\x01")
    for s in range(0, max(len(data), 1), 4096):
        chunk = data[s:s + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) <= 4096 and len(comp) + 2 <= 4098:
            res += struct.pack("<H", 0xB000 | (len(comp) + 2 - 3)) + comp
        elif len(chunk) == 4096:                       # 압축이 안 되면 원문 그대로 (4096바이트 덩어리만 허용)
            res += struct.pack("<H", 0x3000 | (4098 - 3)) + chunk
        else:
            raise ValueError("마지막 덩어리를 압축하지 못함")
    return bytes(res)

def ovba_decompress(buf):                            # 자체 확인용
    assert buf[0] == 1; out = bytearray(); p = 1
    while p < len(buf):
        hdr = struct.unpack_from("<H", buf, p)[0]; size = (hdr & 0x0FFF) + 3; comp = hdr & 0x8000
        body = buf[p + 2:p + size]; p += size; start = len(out)
        if not comp: out += body; continue
        q = 0
        while q < len(body):
            flag = body[q]; q += 1
            for bit in range(8):
                if q >= len(body): break
                if flag & (1 << bit):
                    tok = struct.unpack_from("<H", body, q)[0]; q += 2
                    bits, _ = _copy_token_help(len(out) - start)
                    off = (tok >> (16 - bits)) + 1; ln = (tok & (0xFFFF >> bits)) + 3
                    for _ in range(ln): out.append(out[-off])
                else:
                    out.append(body[q]); q += 1
    return bytes(out)


# ---------------------------------------------------------------- 프로젝트 보호 값 암호화 (2.4.3)
def ovba_encrypt(data, project_id):
    seed = random.randrange(256); version = 2
    key = sum(project_id.encode("ascii")) & 0xFF
    out = bytearray([seed, seed ^ version, seed ^ key])
    unenc1, enc1, enc2 = key, seed ^ key, seed ^ version
    def put(b):
        nonlocal unenc1, enc1, enc2
        e = b ^ ((enc2 + unenc1) & 0xFF); out.append(e)
        enc2, enc1, unenc1 = enc1, e, b
    for _ in range((seed & 6) // 2): put(random.randrange(256))
    for b in struct.pack("<I", len(data)): put(b)
    for b in data: put(b)
    return out.hex().upper()


# ---------------------------------------------------------------- dir 스트림 (2.3.4.2)
def rec(rid, payload):
    return struct.pack("<HI", rid, len(payload)) + payload

def dir_stream(modules):
    b = bytearray()
    b += rec(0x0001, struct.pack("<I", 1))                       # SYSKIND Win32 (64비트 오피스도 읽음)
    b += rec(0x0002, struct.pack("<I", 0x0409))                  # LCID (규격상 0x0409 고정)
    b += rec(0x0014, struct.pack("<I", 0x0409))                  # LCIDINVOKE (규격상 0x0409 고정)
    b += rec(0x0003, struct.pack("<H", CODEPAGE))
    b += rec(0x0004, b"VBAProject")
    b += rec(0x0005, b"") + rec(0x0040, b"")                     # DOCSTRING
    b += rec(0x0006, b"") + rec(0x003D, b"")                     # HELPFILEPATH 1·2
    b += rec(0x0007, struct.pack("<I", 0))                       # HELPCONTEXT
    b += rec(0x0008, struct.pack("<I", 0))                       # LIBFLAGS
    b += struct.pack("<HIIH", 0x0009, 4, 1, 0)                   # VERSION (Reserved=4, Major, Minor)
    b += rec(0x000C, b"") + rec(0x003C, b"")                     # CONSTANTS
    b += rec(0x000F, struct.pack("<H", len(modules)))            # PROJECTMODULES
    b += rec(0x0013, struct.pack("<H", 0xFFFF))                  # PROJECTCOOKIE
    for name, kind, _ in modules:
        nb = name.encode(CP); nu = name.encode("utf-16-le")
        b += rec(0x0019, nb) + rec(0x0047, nu)                   # MODULENAME
        b += rec(0x001A, nb) + rec(0x0032, nu)                   # MODULESTREAMNAME
        b += rec(0x001C, b"") + rec(0x0048, b"")                 # MODULEDOCSTRING
        b += rec(0x0031, struct.pack("<I", 0))                   # MODULEOFFSET (컴파일 캐시 없음 → 소스가 0에서 시작)
        b += rec(0x001E, struct.pack("<I", 0))                   # MODULEHELPCONTEXT
        b += rec(0x002C, struct.pack("<H", 0xFFFF))              # MODULECOOKIE
        b += struct.pack("<HI", 0x0021 if kind == "std" else 0x0022, 0)   # MODULETYPE
        b += struct.pack("<HI", 0x002B, 0)                       # 모듈 끝
    b += struct.pack("<HI", 0x0010, 0)                           # dir 끝
    return bytes(b)


# ---------------------------------------------------------------- 복합 문서(CFB v3) 쓰기
SECT, MINI, CUTOFF = 512, 64, 4096
FREE, ENDCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD

class Node:
    def __init__(self, name, data=None, kids=None):
        self.name, self.data, self.kids = name, data, kids
        self.sid = self.left = self.right = self.child = FREE
        self.start, self.size = ENDCHAIN, 0

def _key(n): return (len(n.name), n.name.upper())

def _tree(nodes):                                         # 균형 이진 트리 (모두 검정)
    if not nodes: return FREE
    nodes = sorted(nodes, key=_key); m = len(nodes) // 2
    nodes[m].left = _tree(nodes[:m]); nodes[m].right = _tree(nodes[m + 1:])
    return nodes[m].sid

def write_cfb(root_kids):
    root = Node("Root Entry", kids=root_kids)
    order = []
    def walk(n):
        n.sid = len(order); order.append(n)
        for k in n.kids or []: walk(k)
    walk(root)
    for n in order:
        if n.kids is not None: n.child = _tree(n.kids)
    # 작은 스트림 → 미니 스트림
    mini = bytearray(); minifat = []
    for n in order:
        if n.data is not None and len(n.data) < CUTOFF:
            n.size = len(n.data)
            if not n.data: n.start = ENDCHAIN; continue
            n.start = len(mini) // MINI; cnt = -(-len(n.data) // MINI)
            minifat += [n.start + k + 1 for k in range(cnt - 1)] + [ENDCHAIN]
            mini += n.data + b"\0" * (cnt * MINI - len(n.data))
    root.data = bytes(mini)
    # 일반 섹터에 들어갈 덩어리: 큰 스트림, 미니 스트림, 미니 FAT, 디렉터리
    blobs = []
    def place(data):
        cnt = max(1, -(-len(data) // SECT)); blobs.append((data, cnt)); return cnt
    big = [n for n in order if n.data is not None and (len(n.data) >= CUTOFF or n is root) and n.data]
    sizes = []
    for n in big: sizes.append(place(n.data))
    mf = b"".join(struct.pack("<I", x) for x in minifat); mf += struct.pack("<I", FREE) * ((-len(mf) // 4) % (SECT // 4))
    if minifat: place(mf)
    dir_count = -(-len(order) // 4); dirdata_len = dir_count * SECT
    total = sum(c for _, c in blobs) + dir_count
    nfat = 1
    while (total + nfat) > nfat * (SECT // 4): nfat += 1
    assert nfat <= 109
    fat = []; pos = 0; starts = []
    for data, cnt in blobs:
        starts.append(pos); fat += [pos + k + 1 for k in range(cnt - 1)] + [ENDCHAIN]; pos += cnt
    dir_start = pos; fat += [pos + k + 1 for k in range(dir_count - 1)] + [ENDCHAIN]; pos += dir_count
    fat_start = pos; fat += [FATSECT] * nfat; pos += nfat
    fat += [FREE] * (nfat * (SECT // 4) - len(fat))
    for n, st in zip(big, starts): n.start = st; n.size = len(n.data)
    if not root.data: root.start, root.size = ENDCHAIN, 0
    minifat_start = starts[len(big)] if minifat else ENDCHAIN
    # 디렉터리 항목
    d = bytearray()
    for n in order:
        nm = n.name.encode("utf-16-le") + b"\0\0"
        typ = 5 if n is root else (1 if n.kids is not None else 2)
        d += nm + b"\0" * (64 - len(nm)) + struct.pack("<HBB", len(nm), typ, 1)
        d += struct.pack("<III", n.left, n.right, n.child) + b"\0" * 16 + b"\0" * 4 + b"\0" * 16
        d += struct.pack("<IQ", n.start if (n.data or n is root) else 0, n.size if n.kids is None or n is root else 0)
    while len(d) < dirdata_len:
        d += b"\0" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", FREE, FREE, FREE) + b"\0" * 36 + struct.pack("<IQ", 0, 0)
    hdr = bytearray(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\0" * 16)
    hdr += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
    hdr += struct.pack("<IIIIIIIII", 0, nfat, dir_start, 0, CUTOFF, minifat_start, len(mf) // SECT if minifat else 0, ENDCHAIN, 0)
    hdr += b"".join(struct.pack("<I", fat_start + k) for k in range(nfat)) + struct.pack("<I", FREE) * (109 - nfat)
    assert len(hdr) == 512
    body = bytearray()
    for data, cnt in blobs: body += data + b"\0" * (cnt * SECT - len(data))
    body += d
    body += b"".join(struct.pack("<I", x) for x in fat)
    return bytes(hdr + body)


# ---------------------------------------------------------------- vbaProject.bin
DOC_ATTR = ('Attribute VB_Name = "{n}"\r\nAttribute VB_Base = "0{{{base}}}"\r\nAttribute VB_GlobalNameSpace = False\r\n'
            'Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n')
WB_BASE, WS_BASE = "00020819-0000-0000-C000-000000000046", "00020820-0000-0000-C000-000000000046"
THISWB_CODE = ("Option Explicit\r\n\r\n' 파일을 열면 설정 시트에 매크로 버튼이 없을 때 만듭니다.\r\n"
               "Private Sub Workbook_Open()\r\n    On Error Resume Next\r\n    EnsureButtons\r\nEnd Sub\r\n")

def vba_project(sheet_codenames, bas_text):
    mods = [("ThisWorkbook", "doc", DOC_ATTR.format(n="ThisWorkbook", base=WB_BASE) + THISWB_CODE)]
    mods += [(c, "doc", DOC_ATTR.format(n=c, base=WS_BASE)) for c in sheet_codenames]
    mods += [("modTrace", "std", bas_text.replace("\r\n", "\n").replace("\n", "\r\n"))]
    pid = "{" + str(uuid.UUID(int=random.getrandbits(128))).upper() + "}"
    proj = [f'ID="{pid}"'] + [f"Document={n}/&H00000000" for n, k, _ in mods if k == "doc"] + ["Module=modTrace",
            'Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
            f'CMG="{ovba_encrypt(bytes(4), pid)}"', f'DPB="{ovba_encrypt(bytes(1), pid)}"', f'GC="{ovba_encrypt(bytes([0xFF]), pid)}"',
            "", "[Host Extender Info]", "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000", "", "[Workspace]"]
    proj += [f"{n}=0, 0, 0, 0, C" for n, _, _ in mods] + [""]
    wm = b"".join(n.encode(CP) + b"\0" + n.encode("utf-16-le") + b"\0\0" for n, _, _ in mods) + b"\0\0"
    streams = [Node(n, ovba_compress(src.encode(CP))) for n, _, src in mods]
    vba = Node("VBA", kids=[Node("dir", ovba_compress(dir_stream(mods))), Node("_VBA_PROJECT", bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]))] + streams)
    return write_cfb([vba, Node("PROJECT", "\r\n".join(proj).encode(CP)), Node("PROJECTwm", wm)]), mods


# ---------------------------------------------------------------- xlsx → xlsm
def main():
    random.seed(20261001)                                      # 같은 입력이면 같은 파일
    zin = zipfile.ZipFile(SRC_XLSX)
    wbx = zin.read("xl/workbook.xml").decode("utf-8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)) | {a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)}
    sheets = re.findall(r'<sheet [^>]*?r:id="(rId\d+)"', wbx)
    files = ["xl/" + rid2target[r].lstrip("/").replace("xl/", "", 1) for r in sheets]
    code = {f: f"Sheet{i + 1}" for i, f in enumerate(files)}
    bas = open(SRC_BAS, encoding=CP).read()
    vbabin, mods = vba_project([code[f] for f in files], bas)
    zout = zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "[Content_Types].xml":
            t = data.decode("utf-8").replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                                              "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
            t = t.replace("</Types>", '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
            data = t.encode("utf-8")
        elif item.filename == "xl/_rels/workbook.xml.rels":
            data = rels.replace("</Relationships>", '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>').encode("utf-8")
        elif item.filename == "xl/workbook.xml":
            t = wbx
            if "<workbookPr" in t: t = re.sub(r"<workbookPr(?=[ />])", '<workbookPr codeName="ThisWorkbook"', t, count=1)
            else: t = re.sub(r"(<workbook [^>]*>)", r'\1<workbookPr codeName="ThisWorkbook"/>', t, count=1)
            data = t.encode("utf-8")
        elif item.filename in code:
            t = data.decode("utf-8"); cn = code[item.filename]
            if "<sheetPr" in t: t = re.sub(r"<sheetPr(?=[ />])", f'<sheetPr codeName="{cn}"', t, count=1)
            else: t = re.sub(r"(<worksheet [^>]*>)", rf'\1<sheetPr codeName="{cn}"/>', t, count=1)
            data = t.encode("utf-8")
        zout.writestr(item, data)
    zout.writestr("xl/vbaProject.bin", vbabin)
    zout.close()
    print("saved", OUT, len(vbabin), "bytes vbaProject ·", len(mods), "modules")


if __name__ == "__main__":
    main()
