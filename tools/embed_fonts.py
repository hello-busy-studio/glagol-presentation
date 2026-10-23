"""Встраивает полные шрифты Wix Madefor Display в готовый .pptx (DESIGN.md §0.3).

Usage:
    python3 tools/embed_fonts.py in.pptx out.pptx [--fonts assets/fonts]

Берёт полные TTF из assets/fonts (OFL, Google Fonts), упаковывает их в Embedded OpenType (.fntdata,
без сжатия и без урезания) и прописывает в presentation.xml. Уже встроенные шрифты (в том числе урезанные
из шаблона) заменяются полными, поэтому скрипт можно запускать и для презентаций, собранных из шаблона.
"""
import os, re, struct, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
FACES = [  # (typeface в pptx, слот, файл)
    ("Wix Madefor Display", "regular", "WixMadeforDisplay-Regular.ttf"),
    ("Wix Madefor Display", "bold", "WixMadeforDisplay-Bold.ttf"),
    ("Wix Madefor Display Medium", "regular", "WixMadeforDisplay-Medium.ttf"),
    ("Wix Madefor Display SemiBold", "regular", "WixMadeforDisplay-SemiBold.ttf"),
    ("Wix Madefor Display ExtraBold", "regular", "WixMadeforDisplay-ExtraBold.ttf"),
]
REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"


def _tables(ttf):
    n = struct.unpack(">H", ttf[4:6])[0]
    t = {}
    for i in range(n):
        tag, _, off, ln = struct.unpack(">4sIII", ttf[12 + 16 * i: 28 + 16 * i])
        t[tag.decode("latin-1")] = ttf[off: off + ln]
    return t


def _name(name_tbl, nid):
    fmt, count, so = struct.unpack(">HHH", name_tbl[:6])
    best = ""
    for i in range(count):
        pid, eid, lid, n, ln, off = struct.unpack(">HHHHHH", name_tbl[6 + 12 * i: 18 + 12 * i])
        if n == nid and pid == 3 and lid in (0x409, 0x419, 0):
            best = name_tbl[so + off: so + off + ln].decode("utf-16-be")
            if lid == 0x409:
                return best
    return best


def eot(ttf, family=None, style=None):
    """Полный несжатый EOT v2.2 (flags = 0) из TTF."""
    t = _tables(ttf)
    os2, head, name = t["OS/2"], t["head"], t["name"]
    weight = struct.unpack(">H", os2[4:6])[0]
    fstype = struct.unpack(">H", os2[8:10])[0]
    panose = os2[32:42]
    ur = struct.unpack(">IIII", os2[42:58])
    cp = struct.unpack(">II", os2[78:86]) if len(os2) >= 86 else (1, 0)
    fssel = struct.unpack(">H", os2[62:64])[0]
    chk = struct.unpack(">I", head[8:12])[0]
    fam = family or _name(name, 1)
    sty = style or _name(name, 2)
    ver = _name(name, 5)
    full = (fam + " " + sty) if family else _name(name, 4)

    def s(x):
        b = x.encode("utf-16-le")
        return struct.pack("<H", len(b)) + b

    body = (panose + bytes([1, 1 if fssel & 1 else 0]) + struct.pack("<IHH", weight, fstype, 0x504C)
            + struct.pack("<IIII", *ur) + struct.pack("<II", *cp) + struct.pack("<I", chk) + b"\0" * 16
            + b"\0\0" + s(fam) + b"\0\0" + s(sty) + b"\0\0" + s(ver) + b"\0\0" + s(full)
            + b"\0\0" + s("")                                   # RootString
            + struct.pack("<II", 0, 0) + b"\0\0" + struct.pack("<H", 0)   # RootStringCheckSum, EUDCCodePage, Padding6, SignatureSize
            + struct.pack("<II", 0, 0))                          # EUDCFlags, EUDCFontSize
    total = 16 + len(body) + len(ttf)
    return struct.pack("<IIII", total, len(ttf), 0x00020002, 0) + body + ttf


def main(src, out, fonts_dir, faces=FACES):
    Z = zipfile.ZipFile(src)
    pres = Z.read("ppt/presentation.xml").decode()
    rels = Z.read("ppt/_rels/presentation.xml.rels").decode()
    ct = Z.read("[Content_Types].xml").decode()

    # убрать старые встроенные шрифты (часто урезанные)
    old = set(re.findall(r'Target="(fonts/[^"]+)"', rels))
    rels = re.sub(r'<Relationship [^>]*Type="%s"[^>]*/>' % re.escape(REL_FONT), "", rels)
    pres = re.sub(r"<p:embeddedFontLst>.*?</p:embeddedFontLst>", "", pres, flags=re.S)
    pres = re.sub(r' saveSubsetFonts="[^"]*"', "", pres)

    used = set(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)); nxt = max(used or {0}) + 100
    files, fams = {}, {}
    for i, (typeface, slot, fn) in enumerate(faces):
        path = os.path.join(fonts_dir, fn)
        sty = "Bold" if slot == "bold" else "Regular"
        data = eot(open(path, "rb").read(), family=typeface, style=sty)
        target = f"fonts/font{i + 1}.fntdata"; rid = f"rId{nxt + i}"
        files["ppt/" + target] = data
        rels = rels.replace("</Relationships>",
                            f'<Relationship Id="{rid}" Type="{REL_FONT}" Target="{target}"/></Relationships>')
        fams.setdefault(typeface, []).append((slot, rid))
    lst = "<p:embeddedFontLst>" + "".join(
        f'<p:embeddedFont><p:font typeface="{tf}"/>' + "".join(f'<p:{sl} r:id="{rid}"/>' for sl, rid in sorted(v, key=lambda a: a[0] != "regular")) + "</p:embeddedFont>"
        for tf, v in fams.items()) + "</p:embeddedFontLst>"
    pres = re.sub(r"(<p:notesSz[^>]*/>)", r"\1" + lst, pres, 1)
    if "embedTrueTypeFonts" not in pres:
        pres = pres.replace("<p:presentation ", '<p:presentation embedTrueTypeFonts="1" ', 1)
    if 'Extension="fntdata"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as O:
        for it in Z.infolist():
            if it.filename.startswith("ppt/fonts/"):
                continue
            data = Z.read(it.filename)
            if it.filename == "ppt/presentation.xml": data = pres.encode()
            elif it.filename == "ppt/_rels/presentation.xml.rels": data = rels.encode()
            elif it.filename == "[Content_Types].xml": data = ct.encode()
            O.writestr(it, data)
        for n, d in files.items():
            O.writestr(n, d)
    print(f"embedded {len(files)} full font files (replaced {len(old)})")


if __name__ == "__main__":
    a = sys.argv[1:]
    fd = os.path.join(HERE, "..", "assets", "fonts")
    if "--fonts" in a:
        i = a.index("--fonts"); fd = a[i + 1]; del a[i:i + 2]
    if len(a) == 3 and a[0].endswith("Шаблон презентации.pptx"):   # старый вызов: шаблон in out
        a = a[1:]
    main(a[0], a[1], fd)
