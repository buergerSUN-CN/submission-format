#!/usr/bin/env python3
"""submission-format 引擎 — 把任意稿件渲染成用户的 clean 投稿 Word 格式。

通用：不写死任何具体稿件的内容/路径/章节名。内容统一从 pandoc AST 抽取，
按 references/target-format-spec.md 的格式规格渲染成全 Normal + 直接格式的 docx。

Usage:
  build_submission.py INPUT -o OUT.docx
        [--resource-path DIR]...   图片/资源查找目录(可多次)
        [--bibliography FILE]      单独的参考文献源(thebibliography/.bib/.md)，附为 References
        [--styles STYLES.docx]     样式资产包(默认 assets/reference_styles.docx)
        [--landscape-mincols N]    >=N 列的表判为宽表→横向页(默认 6)
        [--body-only]              片段模式：跳过 title page/Summary，只渲染正文

INPUT 支持 pandoc 能读的任意格式(.md/.tex/.docx/.html/...)；.pdf 请先转文本。
完整 title page 需要输入含作者/单位/通讯元数据(elsarticle 的 \\author/\\affiliation/\\ead
会被解析；其它格式取 pandoc meta 能给的)。
"""
import argparse, json, os, re, struct, subprocess, sys, glob, shutil, zipfile

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_STYLES = os.path.join(SKILL_DIR, 'assets', 'reference_styles.docx')

TNR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>'
PORTRAIT_SECT = ('<w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="0" w:footer="0" w:gutter="0"/>'
                 '<w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/>')
LANDSCAPE_SECT = ('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
                  '<w:pgMar w:top="1080" w:right="1440" w:bottom="1080" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/>'
                  '<w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/>')
CONTENT_W_PORTRAIT = (11906 - 2*1800) * 635   # EMU
CONTENT_W_LANDSCAPE = (16838 - 2*1440) * 635

def esc(s):
    return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

_MATH = {r'\cdot':'·', r'\times':'×', r'\pm':'±', r'\geq':'≥', r'\leq':'≤', r'\ge':'≥',
         r'\le':'≤', r'\neq':'≠', r'\approx':'≈', r'\rightarrow':'→', r'\to':'→',
         r'\leftarrow':'←', r'\%':'%', r'\,':'', r'\;':' ', r'\alpha':'α', r'\beta':'β',
         r'\chi':'χ', r'\kappa':'κ', r'\mu':'μ', r'\sigma':'σ', r'\circ':'°', r'\degree':'°'}
def math_to_unicode(src):
    s = src
    for k,v in sorted(_MATH.items(), key=lambda kv:-len(kv[0])): s = s.replace(k,v)
    s = re.sub(r'\^\{?([\w.+\-]+)\}?', r'\1', s)
    s = re.sub(r'_\{?([\w.+\-]+)\}?', r'\1', s)
    s = re.sub(r'\\[a-zA-Z]+', '', s)
    return s.replace('{','').replace('}','').replace('$','')

# ───────────────────────── run / para 原语 ─────────────────────────
def run(text, b=False, i=False, sup=False, sz=None):
    rpr = [TNR]
    if b: rpr.append('<w:b/><w:bCs/>')
    if i: rpr.append('<w:i/><w:iCs/>')
    if sup: rpr.append('<w:vertAlign w:val="superscript"/>')
    if sz: rpr.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    return f'<w:r><w:rPr>{"".join(rpr)}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'

def para(runs_xml, style='Normal', jc='both', numId=None, before=None, after=None, sect=None):
    p = [f'<w:pStyle w:val="{style}"/>']
    if numId is not None:
        p.append(f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{numId}"/></w:numPr>')
    sp = []
    if before is not None: sp.append(f'w:before="{before}"')
    if after is not None: sp.append(f'w:after="{after}"')
    if sp: p.append(f'<w:spacing {" ".join(sp)}/>')
    if jc: p.append(f'<w:jc w:val="{jc}"/>')
    if sect: p.append(f'<w:sectPr>{sect}</w:sectPr>')
    return f'<w:p><w:pPr>{"".join(p)}</w:pPr>{runs_xml}</w:p>'

HEAD_RPR = f'{TNR}<w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="32"/>'
def pagebreak():
    return f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:rPr>{HEAD_RPR}</w:rPr></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'
def h1(text, before=None):
    return para(run(text, b=True, sz=28), jc='both', before=before, after=160)
def h2(text):
    return para(run(text, b=True), jc='both')

# ───────────────────────── AST inline → runs ─────────────────────────
def inlines_to_runs(ils, b=False, i=False, sup=False, sz=None):
    out = []
    for n in ils or []:
        t = n.get('t')
        if t == 'Str': out.append(run(n['c'], b, i, sup, sz))
        elif t in ('Space','SoftBreak'): out.append(run(' ', b, i, sup, sz))
        elif t == 'LineBreak': out.append('<w:r><w:br/></w:r>')
        elif t == 'Strong': out.append(inlines_to_runs(n['c'], True, i, sup, sz))
        elif t == 'Emph': out.append(inlines_to_runs(n['c'], b, True, sup, sz))
        elif t == 'Superscript': out.append(inlines_to_runs(n['c'], b, i, True, sz))
        elif t == 'Subscript': out.append(inlines_to_runs(n['c'], b, i, sup, sz))
        elif t == 'Span': out.append(inlines_to_runs(n['c'][1], b, i, sup, sz))
        elif t == 'Quoted':
            q = '“”' if n['c'][0]['t']=='DoubleQuote' else '‘’'
            out.append(run(q[0],b,i,False,sz)+inlines_to_runs(n['c'][1],b,i,sup,sz)+run(q[1],b,i,False,sz))
        elif t in ('Cite','Link'): out.append(inlines_to_runs(n['c'][1], b, i, sup, sz))
        elif t == 'Math': out.append(run(math_to_unicode(n['c'][1]), b, i, sup, sz))
        elif t == 'RawInline': pass
        elif t == 'Note': pass
    return ''.join(out)

def plain_text(ils):
    out = []
    for x in ils or []:
        t = x.get('t')
        if t == 'Str': out.append(x['c'])
        elif t in ('Space','SoftBreak','LineBreak'): out.append(' ')
        elif t == 'Math': out.append(math_to_unicode(x['c'][1]))
        elif t == 'Span': out.append(plain_text(x['c'][1]))
        elif t == 'Quoted': out.append(plain_text(x['c'][1]))
        elif t in ('Strong','Emph','Superscript','Subscript') and isinstance(x.get('c'),list):
            out.append(plain_text(x['c']))
    return ''.join(out)

def caption_runs(kind, n, cap_inlines, sz=18):
    """统一题注：去掉 caption 里已有的 'Table N.'/'Figure N.' 前缀，加粗标签前缀 + 描述。"""
    txt = plain_text(cap_inlines)
    txt = re.sub(rf'^\s*{kind}\s*\d+[.:]?\s*', '', txt, flags=re.I)
    return run(f'{kind} {n}. ', b=True, sz=sz) + run(txt, sz=sz)

def latex_clean(s):
    s = re.sub(r'\\url\{([^}]*)\}', r'\1', s)
    s = re.sub(r'\\href\{[^}]*\}\{([^}]*)\}', r'\1', s)
    s = re.sub(r'\\(emph|textit|textbf|texttt|text)\{([^}]*)\}', r'\2', s)
    s = s.replace('~',' ').replace('--','–').replace(r'\&','&').replace(r'\%','%')
    s = re.sub(r'\\[a-zA-Z]+\*?', '', s).replace('{','').replace('}','')
    return re.sub(r'\s+',' ', s).strip()

def parse_bibitems(path):
    """从 thebibliography/.bbl 自解析 \\bibitem 为有序文献列表(LaTeX 清理后纯文本)。"""
    c = open(path, encoding='utf-8').read()
    m = re.search(r'\\begin\{thebibliography\}(?:\{[^}]*\})?(.*?)\\end\{thebibliography\}', c, re.S)
    body = m.group(1) if m else c
    parts = re.split(r'\\bibitem(?:\[[^\]]*\])?\{[^}]*\}', body)
    return [latex_clean(p) for p in parts[1:] if latex_clean(p)]

# ───────────────────────── frontmatter 抽取 ─────────────────────────
def extract_braced(s, pos):
    depth = 0
    for i in range(pos, len(s)):
        if s[i]=='{': depth += 1
        elif s[i]=='}':
            depth -= 1
            if depth==0: return s[pos+1:i], i+1
    return s[pos+1:], len(s)

def frontmatter_from_tex(tex):
    m = re.search(r'\\title\{', tex)
    title = ''
    if m:
        body,_ = extract_braced(tex, m.end()-1)
        title = re.sub(r'\\textbf\{(.*?)\}', r'\1', body, flags=re.S).strip()
    authors = []
    for m in re.finditer(r'\\(author|ead)(?:\[([^\]]*)\])?\{', tex):
        body,_ = extract_braced(tex, m.end()-1)
        if m.group(1)=='author':
            name = re.sub(r'\\(fnref|corref|thanksref|tnoteref)\{[^}]*\}','',body)
            name = re.sub(r'\\[a-zA-Z]+\*?\s*','',name).strip()
            authors.append({'name':name,'affs':re.findall(r'[\w]+',m.group(2) or ''),
                            'corr':'\\corref' in body,'email':None})
        elif m.group(1)=='ead' and authors and 'url' not in (m.group(2) or ''):
            authors[-1]['email'] = body.strip()
    affs = {}
    for m in re.finditer(r'\\affiliation(?:\[([^\]]*)\])?\{', tex):
        body,_ = extract_braced(tex, m.end()-1)
        org = re.search(r'organization=\{(.*?)\}', body, re.S)
        affs[(m.group(1) or str(len(affs)+1)).strip()] = re.sub(r'\s+',' ', (org.group(1) if org else body)).strip()
    kw = re.search(r'\\begin\{keyword\}(.*?)\\end\{keyword\}', tex, re.S)
    keywords = [re.sub(r'\s+',' ',p).strip() for p in re.split(r'\\sep', kw.group(1))] if kw else []
    keywords = [k for k in keywords if k]
    return title, authors, affs, keywords

def frontmatter_from_meta(meta):
    def mt(x):
        if not x: return ''
        if x['t']=='MetaInlines': return plain_text(x['c'])
        if x['t']=='MetaString': return x['c']
        if x['t']=='MetaList': return [mt(e) for e in x['c']]
        if x['t']=='MetaBlocks': return plain_text(x['c'][0]['c']) if x['c'] else ''
        return ''
    title = mt(meta.get('title'))
    au = meta.get('author'); authors = []
    if au:
        names = mt(au) if au['t']=='MetaList' else [mt(au)]
        authors = [{'name':n,'affs':[],'corr':False,'email':None} for n in (names if isinstance(names,list) else [names])]
    kw = meta.get('keywords') or meta.get('keyword')
    keywords = mt(kw) if kw else []
    if isinstance(keywords,str): keywords = [k.strip() for k in re.split(r'[;,]',keywords) if k.strip()]
    return title, authors, {}, keywords

# ───────────────────────── 图片 ─────────────────────────
def image_size_px(path):
    with open(path,'rb') as f: head = f.read(32)
    if head[:8] == b'\x89PNG\r\n\x1a\n':
        w,h = struct.unpack('>II', head[16:24]); return w,h
    if head[:2] == b'\xff\xd8':  # JPEG
        with open(path,'rb') as f:
            f.read(2)
            while True:
                b = f.read(1)
                while b and b != b'\xff': b = f.read(1)
                marker = f.read(1)
                if not marker: break
                if 0xc0 <= marker[0] <= 0xcf and marker[0] not in (0xc4,0xc8,0xcc):
                    f.read(3); h,w = struct.unpack('>HH', f.read(4)); return w,h
                seg = f.read(2)
                if len(seg)<2: break
                f.seek(struct.unpack('>H',seg)[0]-2, 1)
    return 600,400  # fallback

def find_image(src, resource_paths):
    base, ext = os.path.splitext(src)
    if ext.lower() in ('.pdf','.eps','.svg',''):   # Word 嵌不了矢量/无扩展名 → 优先同名位图
        cands = [base+e for e in ('.png','.jpg','.jpeg')] + ([src] if ext else [])
    else:
        cands = [src, base+'.png', base+'.jpg']
    for rp in ['.'] + resource_paths:
        for c in cands:
            p = c if os.path.isabs(c) else os.path.join(rp, c)
            if os.path.isfile(p): return p
    return None

# ───────────────────────── 表格 ─────────────────────────
def cells_text(row):  # row: list of cells; cell c = [attr, rowspan?, ... , blocks] (pandoc cell)
    out = []
    for cell in row:
        blocks = cell[4] if len(cell) > 4 else cell[-1]
        txt = ' '.join(plain_text(b['c']) for b in blocks if b.get('t') in ('Plain','Para'))
        out.append(txt)
    return out

def render_table(tbl, mincols):
    # tbl c = [attr, caption, colspec, head, bodies, foot]
    colspec = tbl['c'][2]; ncols = len(colspec)
    head = tbl['c'][3]; bodies = tbl['c'][4]
    def rows_of(body_rows):
        return [r['c'][1] if isinstance(r,dict) else r[1] for r in body_rows]
    head_rows = head[1] if head else []
    data_rows = []
    for body in bodies:
        data_rows += body[3]  # body = [attr, rowhdrcols, head, body]
    def row_cells(r):
        cells = r[1] if isinstance(r,list) else r['c'][1]
        return cells
    # 全框线 9pt 表
    border = '<w:tcBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/><w:start w:val="single" w:sz="4" w:color="000000"/><w:end w:val="single" w:sz="4" w:color="000000"/></w:tcBorders>'
    def cell_xml(text, bold, first):
        jc = 'both' if first else 'center'
        r = run(text, b=bold, sz=18)
        ppr = f'<w:pPr><w:pStyle w:val="Normal"/><w:jc w:val="{jc}"/></w:pPr>'
        tcpr = f'<w:tcPr>{border}<w:vAlign w:val="center"/></w:tcPr>'
        return f'<w:tc>{tcpr}<w:p>{ppr}{r}</w:p></w:tc>'
    def cell_text(blocks):
        return ' '.join(plain_text(bl['c']) for bl in blocks if bl.get('t') in ('Plain','Para'))
    def row_xml(cells, bold):
        tcs = ''.join(cell_xml(cell_text(_cell_blocks(c)), bold, j==0) for j,c in enumerate(cells))
        return f'<w:tr>{tcs}</w:tr>'
    # longtable 表头常在 body 重复，去重首行
    if head_rows and data_rows:
        h = [cell_text(_cell_blocks(c)) for c in row_cells(head_rows[-1])]
        d = [cell_text(_cell_blocks(c)) for c in row_cells(data_rows[0])]
        if h == d: data_rows = data_rows[1:]
    rows_xml = []
    for r in head_rows:
        rows_xml.append(row_xml(row_cells(r), bold=True))
    for r in data_rows:
        rows_xml.append(row_xml(row_cells(r), bold=False))
    grid = ''.join('<w:gridCol/>' for _ in range(ncols))
    tbl_xml = (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/>'
               f'<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>'
               f'{"".join(rows_xml)}</w:tbl>')
    is_wide = ncols >= mincols
    return tbl_xml, is_wide

def _cell_blocks(cell):
    # pandoc cell = [attr, alignment, rowspan, colspan, blocks]
    return cell[4] if isinstance(cell,list) and len(cell)>=5 else (cell.get('c',[None,None,None,None,[]])[4] if isinstance(cell,dict) else [])

# ───────────────────────── 正文渲染 ─────────────────────────
NO_BREAK = {'acknowledgements','acknowledgments','author contributions','contributors',
            'conflicts of interest','declaration of interests','competing interests',
            'data availability','data sharing statement','funding','author contribution'}

class Ctx:
    def __init__(s, resource_paths, mincols, pkg):
        s.rp = resource_paths; s.mincols = mincols; s.pkg = pkg
        s.fig_n = 0; s.tab_n = 0; s.rels = []; s.media = []; s.next_rid = 1000; s.in_refs = False; s.ref_n = 0

def add_image(ctx, path):
    ext = os.path.splitext(path)[1].lower().lstrip('.') or 'png'
    ext = 'jpeg' if ext=='jpg' else ext
    rid = f'rIdImg{ctx.next_rid}'; ctx.next_rid += 1
    fname = f'img{ctx.next_rid}.{ext}'
    ctx.media.append((path, fname))
    ctx.rels.append((rid, fname, ext))
    return rid

def figure_xml(ctx, src, caption_blocks):
    path = find_image(src, ctx.rp)
    ctx.fig_n += 1
    out = []
    if path:
        w,h = image_size_px(path)
        cx = min(CONTENT_W_PORTRAIT, int(w*9525))
        cy = int(h * cx / w)
        rid = add_image(ctx, path)
        drawing = (f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" '
            f'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{ctx.fig_n}" name="Figure {ctx.fig_n}"/>'
            f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{ctx.fig_n}" name="Figure {ctx.fig_n}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>')
        out.append(f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr>{TNR}</w:rPr>{drawing}</w:r></w:p>')
    else:
        out.append(para(run(f'[缺图: {src} — 未在 resource-path 找到]', i=True), jc='center'))
    out.append(para(caption_runs('Figure', ctx.fig_n, caption_blocks), jc='both', before=0, after=120))
    return ''.join(out)

def table_block(ctx, tbl):
    ctx.tab_n += 1
    cap_blocks = tbl['c'][1][1]
    cap_inlines = cap_blocks[0]['c'] if cap_blocks else []
    cap_p = para(caption_runs('Table', ctx.tab_n, cap_inlines), jc='both', before=120, after=160)
    tbl_xml, is_wide = render_table(tbl, ctx.mincols)
    if is_wide:
        # 宽表→独占横向节：前段结束纵向节，表后段定义横向节
        pre = para('', sect=PORTRAIT_SECT)
        post = para('', sect=LANDSCAPE_SECT)
        return pre + cap_p + tbl_xml + post
    return cap_p + tbl_xml

def render_blocks(blocks, out, ctx, section_breaks=True, lead_break=True):
    seen_h1 = False
    for b in blocks:
        t = b['t']
        if t == 'Header':
            lvl, _, ils = b['c']; plain = plain_text(ils).strip(); key = plain.lower()
            if lvl == 1:
                ctx.in_refs = key in ('references','bibliography','reference list')
                if ctx.in_refs: ctx.ref_n = 0
                no_lead = not lead_break and not seen_h1   # 片段模式：首个大节不插前导分页符
                seen_h1 = True
                if section_breaks and key not in NO_BREAK and not no_lead:
                    out.append(pagebreak()); out.append(h1(plain))
                else:
                    out.append(h1(plain, before=None if no_lead else 480))
            else:
                out.append(h2(plain))
        elif t in ('Para','Plain'):
            imgs = [n for n in b['c'] if n.get('t')=='Image']
            if len(b['c'])==1 and imgs:
                out.append(figure_xml(ctx, imgs[0]['c'][2][0], imgs[0]['c'][1]))
            elif ctx.in_refs:
                # References：左对齐(jc=start)、自动重编号(去掉 pandoc 残留的开头编号/点)
                ctx.ref_n += 1
                ils = [dict(x) for x in b['c']]
                for x in ils:
                    if x.get('t') in ('Space','SoftBreak'): continue
                    if x.get('t') == 'Str': x['c'] = re.sub(r'^\s*\d*\.?\s*', '', x['c'])
                    break
                out.append(para(run(f'{ctx.ref_n}. ') + inlines_to_runs(ils), jc='start'))
            else:
                out.append(para(inlines_to_runs(b['c']), jc='both'))
        elif t == 'Figure':
            # Figure c=[attr, caption, blocks]
            cap = b['c'][1][1]; cap_inlines = cap[0]['c'] if cap else []
            img = None
            def findimg(x):
                nonlocal img
                if isinstance(x,dict):
                    if x.get('t')=='Image' and img is None: img=x
                    for v in (x.get('c') if isinstance(x.get('c'),list) else []): findimg(v)
                elif isinstance(x,list):
                    for y in x: findimg(y)
            findimg(b['c'][2])
            if img: out.append(figure_xml(ctx, img['c'][2][0], cap_inlines or img['c'][1]))
        elif t == 'Table':
            out.append(table_block(ctx, b))
        elif t == 'Div':
            render_blocks(b['c'][1], out, ctx, section_breaks)
        elif t in ('BulletList','OrderedList'):
            items = b['c'] if t=='BulletList' else b['c'][1]
            for it in items:
                out.append(para(inlines_to_runs(it[0]['c']) if it and it[0].get('t') in ('Plain','Para') else '', jc='both'))

def render_references(items, out):
    # 参考文献：左对齐(jc=start，与正文 both 不同)、无缩进、11pt、手打编号。
    out.append(pagebreak()); out.append(h1('References'))
    for i, txt in enumerate(items, 1):
        out.append(para(run(f'{i}. {txt}'), jc='start'))

# ───────────────────────── title page ─────────────────────────
def title_page(title, authors, affs, keywords, abstract_blocks):
    out = [para(run(title, b=True, sz=28), jc=None)]
    if authors:
        aff_ids = list(affs.keys()) or sorted({a for au in authors for a in au['affs']})
        letter = {aid: chr(ord('a')+i) for i,aid in enumerate(aff_ids)}
        ci=0; corr_mark={}
        for a in authors:
            if a['corr']: ci+=1; corr_mark[id(a)]='*'*ci
        ar=[]
        for idx,a in enumerate(authors):
            ar.append(run(a['name'], sz=22))
            sup=''.join(letter.get(x,'') for x in a['affs'])
            if id(a) in corr_mark: sup += (',' if sup else '')+corr_mark[id(a)]
            if sup: ar.append(run(sup, sz=22, sup=True))
            if idx<len(authors)-1: ar.append(run(', ', sz=22))
        out.append(para(''.join(ar), style='author', jc='both', before=0, after=156))
        for aid in aff_ids:
            if affs.get(aid): out.append(para(run(affs[aid], sz=22), style='address', jc='both', numId=1))
        cnt=0
        for a in authors:
            if a['corr'] and a['email']:
                cnt+=1
                out.append(para(run(f'{"*"*cnt} Corresponding authors: {a["name"]}: {a["email"]};', sz=22),
                                style='address', jc='both', before=156, after=0))
    out.append(pagebreak()); out.append(h1('Summary'))
    for blk in abstract_blocks:
        ils = blk.get('c',[])
        if ils and ils[0].get('t')=='Strong':
            label = plain_text(ils[0]['c']).rstrip('.').strip()
            out.append(para(run(label, b=True), jc='both'))
            rest = ils[1:]
            while rest and rest[0].get('t') in ('Space','SoftBreak'): rest = rest[1:]
            out.append(para(inlines_to_runs(rest), jc='both'))
        else:
            out.append(para(inlines_to_runs(ils), jc='both'))
    if keywords:
        out.append(para(run('Keywords: ', b=True, sz=22)+run('; '.join(keywords)+'.', sz=22),
                        style='address', jc='both', before=156, after=0))
    return out

# ───────────────────────── 打包 ─────────────────────────
def preprocess_latex(tex):
    """array/tabularx 自定义列类型 L/R/C{宽} → p{宽}，让 pandoc 能解析这类表格。"""
    return re.sub(r'(?<![\\A-Za-z])[LRC]\{([^{}]*(?:cm|mm|pt|in|em|\\linewidth|\\textwidth|\\columnwidth)[^{}]*)\}',
                  r'p{\1}', tex)

def run_pandoc_json(path, resource_paths):
    tmp = None; cmd_path = path
    if path.lower().endswith('.tex'):
        tex = preprocess_latex(open(path, encoding='utf-8').read())
        tmp = os.path.join(os.path.dirname(os.path.abspath(path)), '._subfmt_pre.tex')  # 同目录保 \input 基准
        open(tmp, 'w', encoding='utf-8').write(tex); cmd_path = tmp
    cmd = ['pandoc', cmd_path, '-t', 'json']
    if resource_paths: cmd += ['--resource-path', os.pathsep.join(resource_paths)]
    try:
        return json.loads(subprocess.run(cmd, capture_output=True, text=True, check=True).stdout)
    finally:
        if tmp and os.path.exists(tmp): os.remove(tmp)

def repack(pkg_dir, out_path):
    if os.path.exists(out_path): os.remove(out_path)
    with zipfile.ZipFile(out_path,'w',zipfile.ZIP_DEFLATED) as z:
        for root,_,files in os.walk(pkg_dir):
            for fn in files:
                fp = os.path.join(root,fn)
                z.write(fp, os.path.relpath(fp, pkg_dir))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('input'); ap.add_argument('-o','--output', required=True)
    ap.add_argument('--resource-path', action='append', default=[])
    ap.add_argument('--bibliography')
    ap.add_argument('--styles', default=DEFAULT_STYLES)
    ap.add_argument('--landscape-mincols', type=int, default=6)
    ap.add_argument('--body-only', action='store_true',
                    help='片段模式：跳过 title page/Summary，只渲染正文(首个大节不插前导分页符)')
    a = ap.parse_args()
    rp = [os.path.abspath(p) for p in a.resource_path]

    ast = run_pandoc_json(a.input, rp)
    if a.input.lower().endswith('.tex'):
        title, authors, affs, keywords = frontmatter_from_tex(open(a.input,encoding='utf-8').read())
    else:
        title, authors, affs, keywords = frontmatter_from_meta(ast.get('meta',{}))
    if not title:
        title,_,_,_ = frontmatter_from_meta(ast.get('meta',{}))
    abstract = ast.get('meta',{}).get('abstract',{}).get('c',[])

    # 解包样式资产包
    pkg = '/tmp/_subfmt_pkg'; shutil.rmtree(pkg, ignore_errors=True); os.makedirs(pkg)
    with zipfile.ZipFile(a.styles) as z: z.extractall(pkg)
    ctx = Ctx(rp, a.landscape_mincols, pkg)

    body = [] if a.body_only else title_page(title, authors, affs, keywords, abstract)
    render_blocks(ast['blocks'], body, ctx, lead_break=not a.body_only)
    if a.bibliography:
        if a.bibliography.lower().endswith(('.tex','.bbl')):
            items = parse_bibitems(a.bibliography)
        else:
            refs = run_pandoc_json(a.bibliography, rp)
            items = [plain_text(b['c']) for b in refs['blocks'] if b.get('t') in ('Para','Plain')]
        render_references(items, body)

    # 写图 media + rels + content types
    os.makedirs(f'{pkg}/word/media', exist_ok=True)
    for srcpath, fname in ctx.media:
        shutil.copy(srcpath, f'{pkg}/word/media/{fname}')
    rels_path = f'{pkg}/word/_rels/document.xml.rels'
    rels = open(rels_path).read()
    add_rels = ''.join(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{fn}"/>' for rid,fn,_ in ctx.rels)
    rels = rels.replace('</Relationships>', add_rels+'</Relationships>')
    open(rels_path,'w').write(rels)
    ct_path = f'{pkg}/[Content_Types].xml'; ct = open(ct_path).read()
    for ext in {e for _,_,e in ctx.rels}:
        if f'Extension="{ext}"' not in ct:
            ct = ct.replace('</Types>', f'<Default Extension="{ext}" ContentType="image/{ext}"/></Types>')
    open(ct_path,'w').write(ct)

    ns = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {ns}>'
                f'<w:body>{"".join(body)}<w:sectPr>{PORTRAIT_SECT}</w:sectPr></w:body></w:document>')
    open(f'{pkg}/word/document.xml','w',encoding='utf-8').write(document)
    repack(pkg, a.output)
    print(f"[ok] {a.output}  (authors={len(authors)}, figs={ctx.fig_n}, tables={ctx.tab_n})")

if __name__ == '__main__':
    main()
