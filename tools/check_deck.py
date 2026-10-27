"""Автопроверка готовой презентации по DESIGN.md (блок «Перед сдачей обязательно», §10).

Usage:
    python3 tools/check_deck.py deck.pptx [--template "Шаблон презентации.pptx"]

Читает .pptx как zip (нужен только Python 3; Pillow — для оценки переполнения текста, если установлен).
Печатает отчёт: ОШИБКИ (сдавать нельзя) и ПРЕДУПРЕЖДЕНИЯ (проверить глазами на рендере).
Код выхода 1, если есть хотя бы одна ошибка. Проверка не заменяет просмотр PDF-рендера (§10).
"""
import collections, math, os, re, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
EMU = 914400
W, H = 9144000, 5143500
TOL = int(0.02 * EMU)

# §2: все цвета системы. #000000 запрещён отдельно.
PALETTE = {
    "1A1A1A", "F6F6F6", "727272", "FFFFFF", "FF5C00", "C3E5AE", "004139", "DBFB6A", "E0E0E0", "FBA97B",  # тема
    "F8CFB8", "FAB894", "FBA16F", "FC8A4A", "FE7325",                                                    # рыжая шкала
    "E5EBF5", "5E8EDD", "F3F3F3", "D9D9D9", "C6C6C6", "FFBE99",                                          # служебные
    "EFEFEF", "F4CCCC", "444746", "D0D0D0", "C8C8C8",                                                    # из шаблона
}
SCALE = {6, 7, 8, 9, 10, 11, 12, 16, 20, 22, 30, 37, 48, 60, 96}   # §3.2 (+60 — эмодзи §5.9)
FONTS = {"Wix Madefor Display", "Wix Madefor Display Medium", "Wix Madefor Display SemiBold",
         "Wix Madefor Display ExtraBold"}
FILLERS = [  # тексты-заглушки шаблона (§10, п. 8)
    "Заголовок в две строки", "Описание под заголовком", "название слайда", "тут qr", "Джейсон Стейтем",
    "Стрелка вырезана", "Несколько пунктов", "Щёлкните, чтобы", "Click to add", "Lorem ipsum",
]
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def slide_order(Z):
    pres = Z.read("ppt/presentation.xml").decode()
    rels = Z.read("ppt/_rels/presentation.xml.rels").decode()
    tgt = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
    tgt.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
    ids = re.findall(r'<p:sldId [^>]*r:id="([^"]+)"', pres)
    return ["ppt/" + tgt[i].lstrip("/").replace("ppt/", "") for i in ids]


def texts(xml):
    return [re.sub(r"<[^>]+>", "", p) for p in re.findall(r"<a:p>.*?</a:p>|<a:p .*?</a:p>", xml, re.S)]


def top_shapes(xml):
    """Фигуры верхнего уровня spTree: (тег, xml, x, y, cx, cy)."""
    tree = re.search(r"<p:spTree>(.*)</p:spTree>", xml, re.S)
    if not tree:
        return []
    body, out, i = tree.group(1), [], 0
    for m in re.finditer(r"<p:(sp|pic|grpSp|graphicFrame|cxnSp)[ >]", body):
        if m.start() < i:
            continue
        tag = m.group(1)
        depth, j = 0, m.start()
        for t in re.finditer(r"<(/?)p:%s[ >/]" % tag, body[m.start():]):
            depth += -1 if t.group(1) else 1
            if depth == 0:
                j = m.start() + t.end()
                break
        end = body.find(">", j) + 1
        s = body[m.start():end]
        i = end
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"', s)
        x, y, cx, cy = map(int, off.groups()) if off else (0, 0, 0, 0)
        out.append((tag, s, x, y, cx, cy))
    return out


class Report:
    def __init__(self):
        self.err = collections.OrderedDict()
        self.warn = collections.OrderedDict()

    def add(self, kind, key, slide=None):
        d = (self.err if kind == "E" else self.warn).setdefault(key, [])
        if slide is not None and slide not in d:
            d.append(slide)

    def show(self):
        def fmt(sl):
            return "" if not sl else " — слайды " + ", ".join(map(str, sl[:40])) + (" …" if len(sl) > 40 else "")
        print("ОШИБКИ (исправить до сдачи):" if self.err else "Ошибок нет.")
        for k, v in self.err.items():
            print(f"  ✗ {k}{fmt(v)}")
        if self.warn:
            print("ПРЕДУПРЕЖДЕНИЯ (проверить на рендере):")
            for k, v in self.warn.items():
                print(f"  ! {k}{fmt(v)}")


def text_fit(R, n, s, fonts_dir):
    """Грубая оценка: текст не помещается в высоту фигуры (§6.5.1)."""
    try:
        from PIL import ImageFont
    except ImportError:
        return
    if 'wrap="none"' in s or "<a:spAutoFit/>" in s or "<a:normAutofit" in s:
        return
    off = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', s)
    if not off:
        return
    w_pt, h_pt = int(off.group(1)) / 12700, int(off.group(2)) / 12700
    bp = re.search(r"<a:bodyPr([^>]*)", s).group(1) if "<a:bodyPr" in s else ""
    ins = lambda k, d: int(re.search(k + r'="(\d+)"', bp).group(1)) / 12700 if re.search(k + r'="(\d+)"', bp) else d
    w_pt -= ins("lIns", 7.2) + ins("rIns", 7.2)  # по высоте сравниваем с полной плашкой: залезть в отступы не страшно
    if w_pt <= 0:
        return
    paras = re.findall(r"<a:p>.*?</a:p>|<a:p .*?</a:p>", s, re.S)
    while paras and not "".join(texts(paras[-1])).strip():
        paras.pop()  # пустые абзацы в конце не видны
    total, n_lines = 0.0, 0
    for p in paras:
        t = re.sub(r"<[^>]+>", "", re.sub(r"<a:br>.*?</a:br>|<a:br/>", "\n", p, flags=re.S))
        sz = re.search(r' sz="(\d+)"', p)
        size = int(sz.group(1)) / 100 if sz else 18
        tf = re.search(r'<a:latin typeface="([^"]+)"', p)
        face = (tf.group(1) if tf else "Wix Madefor Display").replace("Wix Madefor Display", "").strip() or "Regular"
        path = os.path.join(fonts_dir, f"WixMadeforDisplay-{face}.ttf")
        if not os.path.exists(path):
            return
        f = ImageFont.truetype(path, 100)
        spc = re.search(r'<a:lnSpc><a:spcPct val="(\d+)"', p)
        lh = size * 1.2 * (int(spc.group(1)) / 100000 if spc else 1)
        lines = 0
        for chunk in (t.split("\n") or [""]):
            cur, n_l = "", 1
            for word in chunk.split(" "):
                trial = (cur + " " + word).strip()
                if f.getlength(trial) * size / 100 > w_pt and cur:
                    n_l += 1
                    cur = word
                else:
                    cur = trial
            lines += n_l
        total += lines * lh
        n_lines += lines
    if n_lines > 1 and total > h_pt * 1.1 + 2:
        R.add("W", "текст, похоже, не помещается в свою плашку по высоте (§6.5.1)", n)


def check(path, template=None, fonts_dir=None):
    fonts_dir = fonts_dir or os.path.join(HERE, "..", "assets", "fonts")
    Z = zipfile.ZipFile(path)
    R = Report()
    names = set(Z.namelist())
    pres = Z.read("ppt/presentation.xml").decode()

    # 1. Размер слайда
    m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
    sw, sh = (int(m.group(1)), int(m.group(2))) if m else (W, H)
    if (sw, sh) != (W, H):
        R.add("E", f"размер слайда {sw / EMU:.3f} × {sh / EMU:.3f} in, нужен 10 × 5.625 (§0.1)")

    # 2. Встроенные шрифты
    fnt = [n for n in names if n.startswith("ppt/fonts/")]
    if "<p:embeddedFontLst>" not in pres or not fnt:
        R.add("E", "шрифты не встроены — запусти tools/embed_fonts.py (§0.3)")
    else:
        small = [n for n in fnt if Z.getinfo(n).file_size < 60000]
        if small or 'saveSubsetFonts="1"' in pres:
            R.add("E", "встроены урезанные шрифты — перевстрой полные через tools/embed_fonts.py (§0.3)")

    tpl_texts = set()
    if template and os.path.exists(template):
        T = zipfile.ZipFile(template)
        for n in T.namelist():
            if re.match(r"ppt/slides/slide\d+\.xml$", n):
                tpl_texts |= {t.strip() for t in texts(T.read(n).decode()) if len(t.strip()) >= 25}

    order = slide_order(Z)
    notes = 0
    for i, sp in enumerate(order, 1):
        xml = Z.read(sp).decode()
        rels_name = sp.replace("slides/", "slides/_rels/") + ".rels"
        rels = Z.read(rels_name).decode() if rels_name in names else ""
        hidden = 'show="0"' in xml[:400]

        # Заметки
        nm = re.search(r'Target="\.\./notesSlides/([^"]+)"', rels)
        if nm and "".join(texts(Z.read("ppt/notesSlides/" + nm.group(1)).decode())).strip(" 0123456789"):
            notes += 1

        # Шрифты
        for t in set(re.findall(r'<a:(?:latin|ea|cs) typeface="([^"]+)"', xml)):
            if t not in FONTS and not t.startswith("+") and t != "Apple Color Emoji":
                R.add("E", f"чужой шрифт «{t}» — только Wix Madefor Display (§3.1)", i)
        # Флаг bold: допустим только в задании (20 pt рыжий) и в заголовке фреймворка (8 pt)
        for rpr in re.findall(r"<a:rPr[^>]*\bb=\"1\"[^>]*>", xml):
            sz = re.search(r' sz="(\d+)"', rpr)
            if not sz or int(sz.group(1)) not in (2000, 800):
                R.add("E", "флаг bold — жирность задаётся гарнитурой SemiBold (§3.1)", i)
                break
        # Кегли
        sizes = {int(s) / 100 for s in re.findall(r'<a:(?:rPr|defRPr|endParaRPr)[^>]* sz="(\d+)"', xml)}
        if any(s < 6 for s in sizes):
            R.add("E", "текст мельче 6 pt (§3.2)", i)
        for s in sorted(sizes - SCALE):
            if s >= 6:
                R.add("W", f"кегль {s:g} pt вне шкалы §3.2", i)
        # Цвета
        cols = {c.upper() for c in re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', xml)}
        if "000000" in cols:
            R.add("E", "чистый чёрный #000000 — используй #1A1A1A (§2)", i)
        for c in sorted(cols - PALETTE - {"000000"}):
            R.add("W", f"цвет #{c} вне палитры §2", i)
        if re.search(r'<a:schemeClr val="(?:tx2|bg2)"', xml):
            R.add("W", "цвета tx2/bg2 из схемы — в шаблоне они перепутаны (clrMap), задавай RGB (§0.2)", i)

        # Заглушки и пустота
        all_t = [t.strip() for t in texts(xml) if t.strip()]
        joined = " ".join(all_t)
        for f in FILLERS:
            if f.lower() in joined.lower():
                R.add("E", f"остался текст-заглушка шаблона «{f}…» (§10, п. 8)", i)
        for t in all_t:
            if t in tpl_texts:
                R.add("W", "текст совпадает с текстом шаблона — убедись, что это не заглушка", i)
                break
        has_pic = "<p:pic>" in xml or "<a:blip " in xml or "<p:graphicFrame" in xml
        if not all_t and not has_pic and not hidden:
            R.add("E", "пустой слайд — нет ни текста, ни картинок", i)
        if "<p:ph" in xml:
            empty_ph = [s for s in re.findall(r"<p:sp>.*?</p:sp>", xml, re.S)
                        if "<p:ph" in s and not "".join(texts(s)).strip()]
            if empty_ph:
                R.add("E", "пустой плейсхолдер макета (в режиме правки видно «Щёлкните, чтобы добавить…»)", i)
            else:
                R.add("W", "плейсхолдеры макета вместо свободных фигур на Blank (§0.1)", i)

        # Геометрия
        for tag, s, x, y, cx, cy in top_shapes(xml):
            if cx == 0 and cy == 0:
                continue
            rot = re.search(r'<a:xfrm[^>]*rot="(-?\d+)"', s)
            if rot:  # видимая рамка повёрнутой фигуры
                a = math.radians(int(rot.group(1)) / 60000)
                vw = abs(cx * math.cos(a)) + abs(cy * math.sin(a))
                vh = abs(cx * math.sin(a)) + abs(cy * math.cos(a))
                x, y, cx, cy = x + (cx - vw) / 2, y + (cy - vh) / 2, vw, vh
            if x < -TOL or y < -TOL or x + cx > sw + TOL or y + cy > sh + TOL:
                if cx >= sw * 0.9 or cy >= sh * 0.9:
                    continue  # фон или фото в край
                if tag == "pic":
                    R.add("W", "картинка выходит за край слайда — допустимо только для мокапов и фото в край", i)
                else:
                    R.add("E", "фигура или текст выходит за край слайда", i)
        for tag, s, x, y, cx, cy in top_shapes(xml):
            if tag == "pic":
                big = cx * cy > 0.12 * sw * sh
                edge = x <= TOL or y <= TOL or x + cx >= sw - TOL or y + cy >= sh - TOL  # фото в край
                geom = re.search(r'<a:prstGeom prst="(\w+)"', s)
                if big and not edge and (not geom or geom.group(1) == "rect"):
                    R.add("W", "крупная картинка без скругления углов — если это скриншот, нужен roundRect (§5.11)", i)
                    break
        if "<a:tbl>" not in xml and len(re.findall(r'<p:sp>', xml)) > 30:
            R.add("W", "много мелких фигур — если это таблица, сделай нативную (§6.15)", i)

        # Переполнение текста
        for s in re.findall(r"<p:sp>.*?</p:sp>", xml, re.S):
            if "<p:txBody>" in s and "".join(texts(s)).strip():
                text_fit(R, i, s, fonts_dir)

    if order and notes == 0:
        R.add("W", "нет ни одной заметки спикера — проверь, были ли они в исходнике (сценарий Б)")
    print(f"{os.path.basename(path)}: {len(order)} слайдов")
    R.show()
    return 1 if R.err else 0


if __name__ == "__main__":
    a = sys.argv[1:]
    tpl = None
    if "--template" in a:
        k = a.index("--template"); tpl = a[k + 1]; del a[k:k + 2]
    elif os.path.exists(os.path.join(HERE, "..", "Шаблон презентации.pptx")):
        tpl = os.path.join(HERE, "..", "Шаблон презентации.pptx")
    if not a:
        sys.exit(__doc__)
    sys.exit(check(a[0], tpl))
