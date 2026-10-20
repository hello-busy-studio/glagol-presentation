"""Usage: python3 tools/embed_fonts.py "Шаблон презентации.pptx" in.pptx out.pptx
Copy embedded Wix Madefor Display fonts from the template into a pptx (fntdata parts + embeddedFontLst)."""
import sys, zipfile, re, shutil
tpl, src, out = sys.argv[1:4]
T = zipfile.ZipFile(tpl); Z = zipfile.ZipFile(src)
tpres = T.read('ppt/presentation.xml').decode(); trels = T.read('ppt/_rels/presentation.xml.rels').decode()
lst = re.search(r'<p:embeddedFontLst>.*?</p:embeddedFontLst>', tpres, re.S).group(0)
rids = re.findall(r'r:id="(rId\d+)"', lst)
pres = Z.read('ppt/presentation.xml').decode(); rels = Z.read('ppt/_rels/presentation.xml.rels').decode()
ct = Z.read('[Content_Types].xml').decode()
used = set(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)); nxt = max(used) + 100
files = {}
for rid in rids:
    m = re.search(r'<Relationship[^>]*Id="%s"[^>]*/>' % rid, trels).group(0)
    target = re.search(r'Target="([^"]+)"', m).group(1)
    new = 'rId%d' % nxt; nxt += 1
    lst = lst.replace('r:id="%s"' % rid, 'r:id="%s"' % new)
    rels = rels.replace('</Relationships>', '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="%s"/></Relationships>' % (new, target))
    files['ppt/' + target] = T.read('ppt/' + target)
# embeddedFontLst goes right after notesSz per schema order
pres = re.sub(r'(<p:notesSz[^>]*/>)', r'\1' + lst, pres, 1)
if 'embedTrueTypeFonts' not in pres:
    pres = pres.replace('<p:presentation ', '<p:presentation embedTrueTypeFonts="1" ', 1)
if 'fntdata' not in ct:
    ct = ct.replace('<Default ', '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as O:
    for it in Z.infolist():
        data = Z.read(it.filename)
        if it.filename == 'ppt/presentation.xml': data = pres.encode()
        elif it.filename == 'ppt/_rels/presentation.xml.rels': data = rels.encode()
        elif it.filename == '[Content_Types].xml': data = ct.encode()
        O.writestr(it, data)
    for n, d in files.items(): O.writestr(n, d)
print('embedded', len(files), 'font files')
