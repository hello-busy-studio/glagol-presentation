"""Система отступов (DESIGN.md §4.1): геометрия «на глаз» по метрикам шрифта, без рендера.

Используется в tools/check_deck.py. Можно запускать и отдельно — печатает отступы каждой карточки:
    python3 tools/spacing.py deck.pptx [номер слайда]

Модель текста PowerPoint (замерено на рендере, Wix Madefor Display):
  шаг строки = 1.2 × кегль × интервал;
  первая базовая линия от верха текста = (0.95 − 1.1 × (1 − интервал)) × кегль  (0.95 при 1,0; 0.73 при 0,8);
  высота прописной = 0.715 × кегль.
Визуальный верх текста — верх прописных первой строки, визуальный низ — базовая линия последней строки.
"""
import os, re

EMU = 914400
CAP = 0.715
HERE = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(HERE, "..", "assets", "fonts")
_FONTS = {}


def _font(face):
    if face not in _FONTS:
        from PIL import ImageFont
        p = os.path.join(FONTS_DIR, f"WixMadeforDisplay-{face}.ttf")
        _FONTS[face] = ImageFont.truetype(p, 100) if os.path.exists(p) else None
    return _FONTS[face]


def _texts(xml):
    return [re.sub(r"<[^>]+>", "", p) for p in re.findall(r"<a:p>.*?</a:p>|<a:p .*?</a:p>", xml, re.S)]


def _lines(text, face, size, width_pt, widths=None):
    f = _font(face) or _font("Regular")
    if f is None:
        return 1
    n = 0
    widths = widths if widths is not None else []
    for chunk in text.split("\n"):
        cur, k = "", 1
        # точки переноса: пробел и место после дефиса (неразрывный пробел не рвётся)
        toks = re.findall(r"[^ -]*-|[^ -]+| ", chunk)
        words, w = [], ""
        for t in toks:
            if t == " ":
                words.append((w, " ")); w = ""
            elif t.endswith("-") and len(t) > 1:
                words.append((w + t, "")); w = ""
            else:
                w += t
        words.append((w, ""))
        for word, sep in words:
            trial = cur + word
            if cur.strip() and f.getlength(trial.rstrip()) * size / 100 > width_pt:
                widths.append(f.getlength(cur.rstrip()) * size / 100 * 0.98)
                k += 1
                cur = word + sep
            else:
                cur = trial + sep
        widths.append(f.getlength(cur.rstrip()) * size / 100 * 0.98)
        n += k
    return n


def text_box(s, x, y, cx, cy):
    """Визуальная рамка текста фигуры: dict(l, t, b, size, lines) в дюймах или None."""
    if "<p:txBody>" not in s or not "".join(_texts(s)).strip():
        return None
    bp = re.search(r"<a:bodyPr([^>]*)", s)
    bp = bp.group(1) if bp else ""
    ins = lambda k, d: int(re.search(k + r'="(\d+)"', bp).group(1)) / 12700 if re.search(k + r'="(\d+)"', bp) else d
    l_in, r_in, t_in, b_in = ins("lIns", 7.2), ins("rIns", 7.2), ins("tIns", 3.6), ins("bIns", 3.6)
    width = cx / 12700 - l_in - r_in
    if 'wrap="none"' in bp:
        width = 1e9
    paras = re.findall(r"<a:p>.*?</a:p>|<a:p .*?</a:p>", s, re.S)
    while paras and not "".join(_texts(paras[-1])).strip():
        paras.pop()
    rows = []                               # (кегль, интервал) для каждой строки
    widths = []
    for p in paras:
        t = re.sub(r"<[^>]+>", "", re.sub(r"<a:br>.*?</a:br>|<a:br/>", "\n", p, flags=re.S))
        sz = re.search(r'<a:rPr[^>]* sz="(\d+)"', p) or re.search(r' sz="(\d+)"', p)
        size = int(sz.group(1)) / 100 if sz else 14   # defaultTextStyle шаблона
        tf = re.search(r'<a:latin typeface="([^"]+)"', p)
        face = (tf.group(1) if tf else "").replace("Wix Madefor Display", "").strip() or "Regular"
        spc = re.search(r'<a:lnSpc><a:spcPct val="(\d+)"', p)
        spc = int(spc.group(1)) / 100000 if spc else 1.0
        rows += [(size, spc)] * _lines(t, face, size, max(width, 1), widths)
    if not rows:
        return None
    s0, p0 = rows[0]
    first = (0.95 - 1.1 * (1 - p0)) * s0
    content = sum(1.2 * a * b for a, b in rows)
    last = first + sum(1.2 * a * b for a, b in rows[1:])
    anchor = re.search(r'anchor="(\w+)"', bp)
    anchor = anchor.group(1) if anchor else "t"
    h_pt = cy / 12700
    top = t_in
    if anchor == "ctr":
        top = t_in + (h_pt - t_in - b_in - content) / 2
    elif anchor == "b":
        top = h_pt - b_in - content
    algn = re.search(r'<a:pPr[^>]*algn="(\w+)"', s)
    left = x / EMU + l_in / 72
    if algn and algn.group(1) in ("ctr", "r"):
        f = _font("Regular")
        wmax = max((f.getlength(t) * s0 / 100 for t in _texts(s) if t.strip()), default=width) if f else width
        wmax = min(wmax, width)
        left = x / EMU + (l_in + (width - wmax) / (2 if algn.group(1) == "ctr" else 1)) / 72
    right = left + min(max(widths, default=width), width) / 72
    return dict(l=left, r=right, t=y / EMU + (top + first - CAP * s0) / 72, b=y / EMU + (top + last) / 72,
                size=max(a for a, _ in rows), lines=len(rows))


def _shapes(xml):
    """Фигуры верхнего уровня (группы раскрываются): (tag, xml, x, y, cx, cy) в EMU."""
    tree = re.search(r"<p:spTree>(.*)</p:spTree>", xml, re.S)
    return [sh + (leaf,) for leaf, sh in enumerate(_walk(tree.group(1), (0, 0, 1, 1)))] if tree else []


def _walk(body, tr, top=None):
    out, i, k = [], 0, -1
    ox, oy, kx, ky = tr
    for m in re.finditer(r"<p:(sp|pic|grpSp|graphicFrame|cxnSp)[ >]", body):
        if m.start() < i:
            continue
        tag, depth, j = m.group(1), 0, m.start()
        for t in re.finditer(r"<(/?)p:%s[ >/]" % tag, body[m.start():]):
            depth += -1 if t.group(1) else 1
            if depth == 0:
                j = m.start() + t.end()
                break
        s, i = body[m.start():j], j
        k += 1
        idx = k if top is None else top
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"', s)
        if not off:
            continue
        x, y, cx, cy = map(int, off.groups())
        X, Y, CX, CY = ox + x * kx, oy + y * ky, cx * kx, cy * ky
        if tag == "grpSp":
            ch = re.search(r'<a:chOff x="(-?\d+)" y="(-?\d+)"/>\s*<a:chExt cx="(\d+)" cy="(\d+)"', s)
            inner = s[s.find("</p:grpSpPr>") + 12:] if "</p:grpSpPr>" in s else s
            if ch:
                a, b, c, d = map(int, ch.groups())
                sx, sy = (CX / c if c else 1), (CY / d if d else 1)
                out += _walk(inner, (X - a * sx, Y - b * sy, sx, sy), idx)
            continue
        out.append((tag, s, X, Y, CX, CY, idx))
    return out


def _is_card(tag, s, x, y, cx, cy):
    if tag != "sp" or 'txBox="1"' in s or "".join(_texts(s)).strip():
        return False
    w, h = cx / EMU, cy / EMU
    if not (1.0 <= w <= 9.2 and 0.6 <= h <= 5.3):
        return False
    geom = re.search(r'<a:prstGeom prst="(\w+)"', s)
    if not geom or geom.group(1) not in ("roundRect", "rect"):
        return False
    sppr = s.split("</p:spPr>")[0]
    fill = re.search(r"<p:spPr.*?(<a:solidFill>.*?</a:solidFill>|<a:noFill/>|<a:gradFill)", sppr, re.S)
    has_fill = bool(fill) and "<a:noFill/>" not in fill.group(1)[:12]
    has_line = bool(re.search(r"<a:ln[^>]*>(?:(?!</a:ln>).)*<a:solidFill>", sppr, re.S))
    return has_fill or has_line


def elements(xml):
    """Все элементы слайда с визуальной рамкой: dict(kind, l, t, r, b, size, src, top)."""
    out = []
    raw = _elements(xml)
    return raw


def _elements(xml):
    out = []
    for tag, s, x, y, cx, cy, idx, leaf in _shapes(xml):
        L, T, R, B = x / EMU, y / EMU, (x + cx) / EMU, (y + cy) / EMU
        if tag == "pic" and L > 8.6 and T > 4.6 and cx / EMU < 0.8 and cy / EMU < 0.35:
            out.append(dict(kind="nav", l=L, t=T, r=R, b=B, size=0, src=s, top=idx, leaf=leaf))   # стрелки навигации
            continue
        if re.search(r'<a:xfrm[^>]*rot="-?[1-9]', s) and _is_card(tag, s, x, y, cx, cy):
            continue  # повёрнутый стикер — не карточка
        if _is_card(tag, s, x, y, cx, cy):
            out.append(dict(kind="card", l=L, t=T, r=R, b=B, size=0, src=s, top=idx, leaf=leaf))
            continue
        if tag == "cxnSp" or (tag == "sp" and re.search(r'prst="line"', s)) or (cx == 0 or cy == 0):
            out.append(dict(kind="line", l=L, t=T, r=R, b=B, size=0, src=s, top=idx, leaf=leaf))
            continue
        tb = text_box(s, x, y, cx, cy) if tag == "sp" else None
        boxy = tag == "sp" and 'txBox="1"' not in s and re.search(r"<p:spPr.*?<a:solidFill>", s.split("</p:spPr>")[0], re.S)
        if tb and not boxy:
            short = len(re.sub(r"\s", "", "".join(_texts(s)))) <= 3
            if short and tb["size"] >= 30:      # крупные декоративные цифры
                out.append(dict(kind="decor", l=tb["l"], t=tb["t"], r=R, b=tb["b"], size=tb["size"], src=s, top=idx, leaf=leaf))
                continue
            out.append(dict(kind="text", short=short, l=tb["l"], t=tb["t"], r=tb["r"], b=tb["b"], size=tb["size"], src=s, top=idx, leaf=leaf))
        elif tag in ("sp", "pic", "graphicFrame"):
            out.append(dict(kind="box", l=L, t=T, r=R, b=B, size=0, src=s, top=idx, leaf=leaf))
    return out


def cards(xml):
    """Карточки с содержимым и отступами: список dict(card, items, left, top, bottom, gaps)."""
    els = elements(xml)
    res = []
    for c in [e for e in els if e["kind"] == "card"]:
        inside = [e for e in els if e is not c and e["kind"] != "card" and
                  e["l"] >= c["l"] - 0.02 and e["r"] <= c["r"] + 0.05 and e["t"] >= c["t"] - 0.02 and e["b"] <= c["b"] + 0.3
                  and (e["t"] + e["b"]) / 2 < c["b"]]
        edge = [e for e in inside if e["b"] > c["b"] + 0.02 and e["t"] < c["b"]]      # «язычок» на нижней кромке
        inside = [e for e in inside if e not in edge]
        content = [e for e in inside if e["kind"] in ("text", "box", "decor") and (e["r"] - e["l"]) < (c["r"] - c["l"]) - 0.04]
        if not any(e["kind"] == "text" for e in content):
            continue
        nav = any(e["kind"] == "nav" and e["l"] < c["r"] and e["t"] < c["b"] and e["r"] > c["l"] for e in els)
        decor = [e for e in content if e["kind"] == "decor"]
        content = [e for e in content if e["kind"] != "decor"]
        if content:   # декоративный знак на одной оси с контентом (иконка-звёздочка) — часть контента
            l0 = min(e["l"] for e in content)
            content += [e for e in decor if abs(e["l"] - l0) < 0.1]
            decor = [e for e in decor if e not in content]
        centered = all('algn="ctr"' in e["src"] for e in content if e["kind"] == "text")
        left = min(e["l"] for e in content) - c["l"]
        top = min(e["t"] for e in content) - c["t"]
        bottom = c["b"] - max(e["b"] for e in content)
        its = sorted(content, key=lambda e: e["t"])
        pinned = bool(edge) or any(b["t"] - a["b"] > 2.5 * left for a, b in zip(its, its[1:]))  # низ задан намеренно
        paired = any(e["kind"] in ("box", "card") and e is not c and not (e["l"] >= c["l"] and e["r"] <= c["r"]) and
                     abs(e["t"] - c["t"]) < 0.06 and abs(e["b"] - c["b"]) < 0.06 and (e["r"] - e["l"]) > 0.8
                     and not _is_card_like(e, c) for e in els)                   # рядом фото/плашка той же высоты
        res.append(dict(card=c, items=content, decor=decor, nav=nav, centered=centered, pinned=pinned, paired=paired, lines=[e for e in inside if e["kind"] == "line"],
                        left=left, top=top, bottom=bottom, gaps=_gaps(content, inside)))
    return res


def min_gap(size):
    """Минимальный зазор между строками соседних блоков: 0.08 in, для мелкого текста — 0.8 × кегль."""
    return min(0.08, 0.8 * size / 72)


def _is_card_like(e, c):
    """Соседняя карточка того же размера — это ряд, а не пара «фото + карточка»."""
    return e["kind"] == "card" and abs((e["r"] - e["l"]) - (c["r"] - c["l"])) < 0.05


def _gaps(content, inside):
    """Зазоры «заголовок → текст под ним» внутри карточки: (верхний, нижний, зазор)."""
    texts = sorted([e for e in content if e["kind"] == "text"], key=lambda e: e["t"])
    out = []
    for a in texts:
        below = [b for b in texts if b is not a and b["t"] > a["t"] + 0.02 and
                 min(a["r"], b["r"]) - max(a["l"], b["l"]) > 0.1]
        if not below:
            continue
        b = min(below, key=lambda e: e["t"])
        if any(l["t"] > a["b"] and l["t"] < b["t"] for l in inside if l["kind"] == "line"):
            continue  # между ними разделитель — это разные секции
        out.append((a, b, b["t"] - a["b"]))
    return out


if __name__ == "__main__":
    import sys, zipfile
    sys.path.insert(0, HERE)
    from check_deck import slide_order
    Z = zipfile.ZipFile(sys.argv[1])
    only = int(sys.argv[2]) if len(sys.argv) > 2 else None
    for n, sp in enumerate(slide_order(Z), 1):
        if only and n != only:
            continue
        for c in cards(Z.read(sp).decode()):
            k = c["card"]
            print(f"{n:3} карточка {k['l']:.2f},{k['t']:.2f} {k['r'] - k['l']:.2f}×{k['b'] - k['t']:.2f}: "
                  f"слева {c['left']:.2f} сверху {c['top']:.2f} снизу {c['bottom']:.2f} "
                  f"зазоры {[round(g, 2) for _, _, g in c['gaps']]}")
