#!/usr/bin/env python3
"""submission-format 引擎 — 把任意稿件渲染成用户的 clean 投稿 Word 格式。

通用：不写死任何具体稿件的内容/路径/章节名。内容统一从 pandoc AST 抽取，
按 references/target-format-spec.md 的格式规格渲染成全 Normal + 直接格式的 docx。

Usage:
  build_submission.py INPUT -o OUT.docx
        [--resource-path DIR]...   图片/资源查找目录(可多次)
        [--bibliography FILE]      单独的参考文献源(thebibliography/.bib/.md)，附为 References
        [--styles STYLES.docx]     样式资产包(默认 assets/reference_styles.docx)
        [--landscape-fit F]        自然总宽 > 纵向页宽×F → 转横向(默认 1.6)
        [--landscape-mincols N]    可选 override：列数 >= N 一律横向(默认 99=关)
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
TBL_DXA_PORTRAIT = 11906 - 2*1800    # 8306  纵向表可用宽(dxa)
TBL_DXA_LANDSCAPE = 16838 - 2*1440   # 13958 横向表可用宽(dxa)
CHAR_DXA = 95        # 9pt 单元格每字符宽度估计(dxa)；列宽分配与横/纵判定共用
CELL_PAD = 216       # 单元格左右内边距合计(start/end=108)
LANDSCAPE_FIT = 1.6  # 自然总宽 > 纵向可用宽 × 此值 → 转横向(否则压进纵向页)；--landscape-fit 可调

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
def run(text, b=False, i=False, sup=False, sz=None, szcs=None, ics=False):
    rpr = [TNR]
    if b: rpr.append('<w:b/><w:bCs/>')
    if i: rpr.append('<w:i/><w:iCs/>')
    elif ics: rpr.append('<w:iCs/>')   # 复杂脚本斜体位(题注/脚注：真值带 iCs 但不带 w:i)
    if sup: rpr.append('<w:vertAlign w:val="superscript"/>')
    if sz: rpr.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{szcs or sz}"/>')
    return f'<w:r><w:rPr>{"".join(rpr)}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'

def para(runs_xml, style='Normal', jc='both', numId=None, before=None, after=None, line=None, sect=None):
    p = [f'<w:pStyle w:val="{style}"/>']
    if numId is not None:
        p.append(f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{numId}"/></w:numPr>')
    sp = []
    if line is not None: sp.append(f'w:lineRule="auto" w:line="{line}"')
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
    return para(run(text, b=True, sz=28, szcs=32), jc='both', before=before, after=160)
def h2(text):
    return para(run(text, b=True), jc='both')

NOTE_SENTINEL = '⟦SUBFMTNOTE⟧'   # \scriptsize/\footnotesize → 此哨兵；标记表脚注(9pt)，渲染前剥除

# ───────────────────────── 正文引用(上标编号) ─────────────────────────
_CITE_KEY2NUM = {}   # bib key → 1-based 编号；main() 在渲染正文前填充(来自 --bibliography 的 \bibitem 顺序)

def _collapse(nums):
    """[1,2,3,5,6] → '1-3,5-6'：连续号用连字符段，跳号用逗号(与真值一致，无空格)。"""
    nums = sorted(set(nums)); out = []; k = 0
    while k < len(nums):
        j = k
        while j+1 < len(nums) and nums[j+1] == nums[j]+1: j += 1
        out.append(str(nums[k]) if j == k else f'{nums[k]}-{nums[j]}')
        k = j+1
    return ','.join(out)

def cite_runs(citations, b=False, i=False, sz=None):
    """pandoc Cite 的 citations 列表(各含 citationId) → 一个上标 run；未知 key 跳过，全空则返回 ''。"""
    nums = [_CITE_KEY2NUM[c['citationId']] for c in (citations or []) if c.get('citationId') in _CITE_KEY2NUM]
    return run(_collapse(nums), b=b, i=i, sup=True, sz=sz) if nums else ''

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
        elif t == 'Cite':
            cr = cite_runs(n['c'][0], b, i, sz)          # 数字上标引用(回退：无映射时取原文内容)
            out.append(cr if cr else inlines_to_runs(n['c'][1], b, i, sup, sz))
        elif t == 'Link': out.append(inlines_to_runs(n['c'][1], b, i, sup, sz))
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

def caption_runs(kind, label, cap_inlines, sz=18):
    """统一题注：剥掉 caption 自带的 'Table N.'/'Table S1.' 前缀(含 S 编号，消双前缀)，
    用传入 label(主体为数字、补充材料为 'S1'…)重排为加粗标签前缀 + 描述。"""
    txt = plain_text(cap_inlines)
    txt = re.sub(rf'^\s*{kind}\s*S?\d+[.:]?\s*', '', txt, flags=re.I)
    return run(f'{kind} {label}. ', b=True, sz=sz) + run(txt, sz=sz)

def latex_clean(s):
    s = re.sub(r'\\url\{([^}]*)\}', r'\1', s)
    s = re.sub(r'\\href\{[^}]*\}\{([^}]*)\}', r'\1', s)
    s = re.sub(r'\\(emph|textit|textbf|texttt|text)\{([^}]*)\}', r'\2', s)
    s = s.replace('~',' ').replace('--','–').replace(r'\&','&').replace(r'\%','%')
    s = re.sub(r'\\[a-zA-Z]+\*?', '', s).replace('{','').replace('}','')
    return re.sub(r'\s+',' ', s).strip()

def parse_bibitems(path):
    """从 thebibliography/.bbl 解析 \\bibitem 为有序 [(key, 纯文本)]；key 供正文上标引用编号。"""
    c = open(path, encoding='utf-8').read()
    m = re.search(r'\\begin\{thebibliography\}(?:\{[^}]*\})?(.*?)\\end\{thebibliography\}', c, re.S)
    body = m.group(1) if m else c
    keys = re.findall(r'\\bibitem(?:\[[^\]]*\])?\{([^}]*)\}', body)
    texts = [latex_clean(p) for p in re.split(r'\\bibitem(?:\[[^\]]*\])?\{[^}]*\}', body)[1:]]
    return [(k, t) for k, t in zip(keys, texts) if t]

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

def _col_nat(col_chars):
    """各列自然宽(dxa) = 最长单元格字符数 × CHAR_DXA + 单元格内边距。"""
    return [max(1, c)*CHAR_DXA + CELL_PAD for c in col_chars]

def _col_widths(col_chars, col_word, total):
    """按各列自然宽成比例分配列宽，使表尽量不换行→整体高度最短。
    - 自然宽合计 <= 可用宽：等比放大，每列 ≥ 自然宽 → 全不换行(高度最小)。
    - 合计 > 可用宽：等比压缩，但每列不低于其"最长单词宽"(避免把数字/单词折断，如 73→7/3)。
    列宽合计 == total。"""
    nat = _col_nat(col_chars)
    mn = [max(480, w*CHAR_DXA + CELL_PAD) for w in col_word]   # 列下限=最长单词不折断
    s = sum(nat) or 1
    w = [max(mn[j], nat[j]*total//s) for j in range(len(nat))]
    diff = total - sum(w)                 # 取整/下限造成的差额
    if w:
        k = w.index(max(w)); w[k] = max(mn[k], w[k] + diff)   # 摊到最宽列
    return w

def _table_colchars(tbl):
    """从 pandoc table 块抽取各列最长单元格字符数(用于横/纵判定与列宽分配)。"""
    ncols = len(tbl['c'][2]); head = tbl['c'][3]; bodies = tbl['c'][4]
    rows = (head[1] if head else []) + [r for body in bodies for r in body[3]]
    col = [1]*ncols
    for rr in rows:
        cells = rr[1] if isinstance(rr,list) else rr['c'][1]
        for j,c in enumerate(cells[:ncols]):
            txt = ' '.join(plain_text(bl['c']) for bl in _cell_blocks(c) if bl.get('t') in ('Plain','Para'))
            col[j] = max(col[j], len(txt))
    return col

def _is_wide(col_chars, ncols, mincols, fit):
    """是否转横向：自然总宽 > 纵向可用宽 × fit(放不进纵向)，或列数 ≥ mincols(手动 override)。"""
    return sum(_col_nat(col_chars)) > TBL_DXA_PORTRAIT * fit or (mincols and ncols >= mincols)

def render_table(tbl, mincols, fit):
    # tbl c = [attr, caption, colspec, head, bodies, foot]
    colspec = tbl['c'][2]; ncols = len(colspec)
    head = tbl['c'][3]; bodies = tbl['c'][4]
    head_rows = head[1] if head else []
    data_rows = []
    for body in bodies:
        data_rows += body[3]  # body = [attr, rowhdrcols, head, body]
    def row_cells(r):
        return r[1] if isinstance(r,list) else r['c'][1]
    def cell_text(blocks):
        return ' '.join(plain_text(bl['c']) for bl in blocks if bl.get('t') in ('Plain','Para'))
    def row_texts(r):
        return [cell_text(_cell_blocks(c)) for c in row_cells(r)]
    head_txt = [row_texts(r) for r in head_rows]
    data_txt = [row_texts(r) for r in data_rows]
    # longtable 表头常在 body 重复，去重首行
    if head_txt and data_txt and head_txt[-1] == data_txt[0]:
        data_txt = data_txt[1:]
    # col_chars[j]=第 j 列最长单元格字符数(自然宽代理)；col_word[j]=最长单词(列下限，防折断)
    col_chars = [1]*ncols; col_word = [1]*ncols
    for row in head_txt + data_txt:
        for j in range(min(ncols, len(row))):
            col_chars[j] = max(col_chars[j], len(row[j]))
            col_word[j] = max(col_word[j], max((len(w) for w in row[j].split()), default=1))
    is_wide = _is_wide(col_chars, ncols, mincols, fit)   # 按内容宽度判定(放不进纵向才转横向)
    total = TBL_DXA_LANDSCAPE if is_wide else TBL_DXA_PORTRAIT
    widths = _col_widths(col_chars, col_word, total)
    # 全框线 9pt 表
    border = '<w:tcBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/><w:start w:val="single" w:sz="4" w:color="000000"/><w:end w:val="single" w:sz="4" w:color="000000"/></w:tcBorders>'
    def cell_xml(text, bold, j):
        jc = 'both' if j==0 else 'center'
        r = run(text, b=bold, sz=18)
        # 单元格段落：单倍行距 line=240、无段前后距(真值表格不继承 Normal 的 278/160)
        ppr = (f'<w:pPr><w:pStyle w:val="Normal"/>'
               f'<w:spacing w:lineRule="auto" w:line="240" w:before="0" w:after="0"/>'
               f'<w:jc w:val="{jc}"/></w:pPr>')
        tcpr = f'<w:tcPr><w:tcW w:w="{widths[j]}" w:type="dxa"/>{border}<w:vAlign w:val="center"/></w:tcPr>'
        return f'<w:tc>{tcpr}<w:p>{ppr}{r}</w:p></w:tc>'
    def row_xml(txts, bold):
        tcs = ''.join(cell_xml(txts[j] if j<len(txts) else '', bold, j) for j in range(ncols))
        return f'<w:tr>{tcs}</w:tr>'
    rows_xml = [row_xml(r, True) for r in head_txt] + [row_xml(r, False) for r in data_txt]
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    cellmar = ('<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:start w:w="108" w:type="dxa"/>'
               '<w:bottom w:w="0" w:type="dxa"/><w:end w:w="108" w:type="dxa"/></w:tblCellMar>')
    tbl_xml = (f'<w:tbl><w:tblPr><w:tblW w:w="{total}" w:type="dxa"/><w:jc w:val="center"/>'
               f'<w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/>{cellmar}</w:tblPr>'
               f'<w:tblGrid>{grid}</w:tblGrid>{"".join(rows_xml)}</w:tbl>')
    return tbl_xml, is_wide

def _cell_blocks(cell):
    # pandoc cell = [attr, alignment, rowspan, colspan, blocks]
    return cell[4] if isinstance(cell,list) and len(cell)>=5 else (cell.get('c',[None,None,None,None,[]])[4] if isinstance(cell,dict) else [])

def _is_wide_block(b, mincols, fit):
    """该块是否为(或包裹着)横向宽表——pandoc 把浮动表包进 Div[Table]，故需穿透 Div。
    判定与 render_table 一致(按内容宽度)，用于让相邻横向宽表续在同一横向节、避免夹出空白纵向页。"""
    if b['t'] == 'Table':
        return _is_wide(_table_colchars(b), len(b['c'][2]), mincols, fit)
    if b['t'] == 'Div': return any(_is_wide_block(k, mincols, fit) for k in b['c'][1])
    return False

# ───────────────────────── 正文渲染 ─────────────────────────
NO_BREAK = {'acknowledgements','acknowledgments','author contributions','contributors',
            'conflicts of interest','declaration of interests','competing interests',
            'data availability','data sharing statement','funding','author contribution'}

class Ctx:
    def __init__(s, resource_paths, mincols, pkg, landscape_fit=LANDSCAPE_FIT):
        s.rp = resource_paths; s.mincols = mincols; s.pkg = pkg; s.landscape_fit = landscape_fit
        s.fig_n = 0; s.tab_n = 0; s.rels = []; s.media = []; s.next_rid = 1000; s.in_refs = False; s.ref_n = 0
        s.supp = False          # 进入补充材料后：图/表改用 S 编号
        s.s_fig_n = 0; s.s_tab_n = 0; s.draw_n = 0   # S 计数器 + 图形对象唯一 id 计数(与题注号解耦)
        s.in_ric = False        # Research in context 区(用 address 样式拆分标题/正文)
        s.supp_idx = None       # 参考文献插入点(首个 Supplementary 大节之前)
        s.landscape_close = None  # 横向表后待收尾的 LANDSCAPE 分节段(跨 Div 递归共享，故放 ctx)

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
    if ctx.supp: ctx.s_fig_n += 1; label = f'S{ctx.s_fig_n}'    # 补充材料：Figure S1, S2…
    else: ctx.fig_n += 1; label = ctx.fig_n
    ctx.draw_n += 1; did = ctx.draw_n                           # 图形对象 id(全局唯一，避免主/补冲突)
    out = []
    if path:
        w,h = image_size_px(path)
        cx = min(CONTENT_W_PORTRAIT, int(w*9525))
        cy = int(h * cx / w)
        rid = add_image(ctx, path)
        drawing = (f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" '
            f'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{did}" name="Figure {label}"/>'
            f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{did}" name="Figure {label}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>')
        out.append(f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr>{TNR}</w:rPr>{drawing}</w:r></w:p>')
    else:
        out.append(para(run(f'[缺图: {src} — 未在 resource-path 找到]', i=True), jc='center'))
    out.append(para(caption_runs('Figure', label, caption_blocks), jc='both', before=0, after=120, line=240))
    return ''.join(out)

def footnote_para(text):
    """表脚注：9pt、iCs(不斜)、两端对齐、before=0 after=120(真值表下注样式)。"""
    return para(run(text, sz=18, ics=True), jc='both', before=0, after=120)

def table_block(ctx, tbl):
    if ctx.supp: ctx.s_tab_n += 1; label = f'S{ctx.s_tab_n}'    # 补充材料：Table S1, S2…
    else: ctx.tab_n += 1; label = ctx.tab_n
    cap_blocks = tbl['c'][1][1]
    cap_inlines = cap_blocks[0]['c'] if cap_blocks else []
    cap_p = para(caption_runs('Table', label, cap_inlines), jc='both', before=120, after=160, line=240)
    tbl_xml, is_wide = render_table(tbl, ctx.mincols, ctx.landscape_fit)
    # 只返回 题注+表(不含分节段)；横向分节由 render_blocks 处理，
    # 以便把紧随的表脚注一并圈进横向节(否则脚注被挤到下一页)。
    return cap_p + tbl_xml, is_wide

def render_blocks(blocks, out, ctx, section_breaks=True, lead_break=True, is_top=True):
    seen_h1 = False
    for b in blocks:
        t = b['t']
        closed_ls = False     # 本块是否刚收尾了横向节(供其后大节省掉多余 pagebreak)
        if ctx.landscape_close is not None:
            # 横向表后：表脚注/空段并入横向节；相邻横向宽表续在同一横向节；遇正文内容才收尾
            if t in ('Para','Plain'):
                txt = plain_text(b['c'])
                if NOTE_SENTINEL in txt[:32]:
                    out.append(footnote_para(txt.replace(NOTE_SENTINEL, '').strip())); continue
                if not txt.strip() and not any(n.get('t')=='Image' for n in b['c']):
                    continue
            if not _is_wide_block(b, ctx.mincols, ctx.landscape_fit):
                out.append(ctx.landscape_close); ctx.landscape_close = None; closed_ls = True
            # 否则：下一块仍是横向宽表 → 不收尾、不另插 pre，续在同一横向节
        if t == 'Header':
            lvl, _, ils = b['c']; plain = plain_text(ils).strip(); key = plain.lower()
            if lvl == 1:
                ctx.in_refs = key in ('references','bibliography','reference list')
                if ctx.in_refs: ctx.ref_n = 0
                ctx.in_ric = (key == 'research in context')
                if key.startswith('supplementary'):
                    if ctx.supp_idx is None: ctx.supp_idx = len(out)  # 参考文献插入点：补充材料之前
                    ctx.supp = True                                   # 此后图/表改用 S 编号
                no_lead = not lead_break and not seen_h1   # 片段模式：首个大节不插前导分页符
                seen_h1 = True
                if section_breaks and key not in NO_BREAK and not no_lead and not closed_ls:
                    out.append(pagebreak()); out.append(h1(plain))
                else:
                    # closed_ls：刚收尾横向节(分节符已换页)，不再叠加 pagebreak，免空白页
                    out.append(h1(plain))   # run-on 标题 before=0(靠前段 after=160 分隔)，真值不用 480
            else:
                out.append(h2(plain))
        elif t in ('Para','Plain'):
            imgs = [n for n in b['c'] if n.get('t')=='Image']
            if len(b['c'])==1 and imgs:
                out.append(figure_xml(ctx, imgs[0]['c'][2][0], imgs[0]['c'][1]))
            elif NOTE_SENTINEL in plain_text(b['c'])[:32]:
                out.append(footnote_para(plain_text(b['c']).replace(NOTE_SENTINEL, '').strip()))  # 表脚注(9pt iCs)
            elif ctx.in_refs:
                # References：左对齐(jc=start)、自动重编号(去掉 pandoc 残留的开头编号/点)
                ctx.ref_n += 1
                ils = [dict(x) for x in b['c']]
                for x in ils:
                    if x.get('t') in ('Space','SoftBreak'): continue
                    if x.get('t') == 'Str': x['c'] = re.sub(r'^\s*\d*\.?\s*', '', x['c'])
                    break
                out.append(para(run(f'{ctx.ref_n}. ') + inlines_to_runs(ils), jc='start'))
            elif ctx.in_ric and b['c'] and b['c'][0].get('t') == 'Strong':
                # Research in context：粗体小标题 + 正文各成一段(address 样式、before=156、正文 1.5x)
                lead = b['c'][0]['c']; rest = b['c'][1:]
                while rest and rest[0].get('t') in ('Space','SoftBreak'): rest = rest[1:]
                out.append(para(inlines_to_runs(lead, b=True, sz=22), style='address', jc='both', before=156, after=0))
                out.append(para(inlines_to_runs(rest, sz=22), style='address', jc='both', before=156, after=0, line=360))
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
            xml, is_wide = table_block(ctx, b)
            if is_wide:
                if ctx.landscape_close is None:                       # 从纵向进入横向才插 pre
                    out.append(para('', sect=PORTRAIT_SECT))          # 收尾前面的纵向节
                out.append(xml)                                       # (已在横向节则直接续上)
                ctx.landscape_close = para('', sect=LANDSCAPE_SECT)   # 推迟到表脚注(若有)之后再收尾横向节
            else:
                out.append(xml)
        elif t == 'Div':
            render_blocks(b['c'][1], out, ctx, section_breaks, is_top=False)
        elif t in ('BulletList','OrderedList'):
            items = b['c'] if t=='BulletList' else b['c'][1]
            for it in items:
                out.append(para(inlines_to_runs(it[0]['c']) if it and it[0].get('t') in ('Plain','Para') else '', jc='both'))
    if is_top and ctx.landscape_close is not None:    # 文末仍有未收尾的横向节(末表无脚注)
        out.append(ctx.landscape_close); ctx.landscape_close = None

def references_block(items):
    # 参考文献块：分页符 + 'References' 标题 + 各条(左对齐 jc=start、无缩进、11pt、手打编号)。
    blk = [pagebreak(), h1('References')]
    for i, it in enumerate(items, 1):
        txt = it[1] if isinstance(it, tuple) else it      # items 可能是 (key,text) 或纯文本
        blk.append(para(run(f'{i}. {txt}'), jc='start'))
    return blk

# ───────────────────────── title page ─────────────────────────
def title_page(title, authors, affs, keywords, abstract_blocks):
    out = [para(run(title, b=True, sz=28, szcs=32), jc=None)]
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
def _strip_resizebox(s):
    r"""去掉 \resizebox{w}{h}{<tabular>} 缩放包裹，保留内部表体(否则 pandoc 解析不到表)。"""
    out = []; i = 0
    while True:
        m = re.search(r'\\resizebox\s*\*?\s*', s[i:])
        if not m: out.append(s[i:]); break
        j = i + m.start(); out.append(s[i:j]); k = j + len(m.group(0))
        for _ in range(2):                       # 跳过 {width}{height} 两个参数组
            while k < len(s) and s[k] != '{': k += 1
            _, k = extract_braced(s, k)
        while k < len(s) and s[k] != '{': k += 1
        inner, k = extract_braced(s, k)          # 第三组 = 表体，解包
        out.append(inner); i = k
    return ''.join(out)

def _expand_multicolumn(s):
    r"""\multicolumn{N}{align}{内容} → 内容 + (N-1) 个空单元格(& )。
    pandoc 常把整行 \multicolumn 分类小标题整体丢弃，展开后能在首列保留文字(对齐真值)。"""
    out = []; i = 0
    while True:
        m = re.search(r'\\multicolumn\s*\{(\d+)\}\s*', s[i:])
        if not m: out.append(s[i:]); break
        n = int(m.group(1)); j = i + m.start(); out.append(s[i:j]); k = i + m.end()
        while k < len(s) and s[k] != '{': k += 1   # 跳过 {align}
        _, k = extract_braced(s, k)
        while k < len(s) and s[k] != '{': k += 1   # {内容}
        content, k = extract_braced(s, k)
        out.append(content + ' &'*(n-1)); i = k
    return ''.join(out)

def preprocess_latex(tex):
    r"""让 pandoc 能解析更多表：① L/R/C{宽}→p{宽}；② sidewaystable→table；③ 去 \resizebox 包裹；
    ④ 展开 \multicolumn(否则整行分类小标题被 pandoc 丢弃)。
    (宽表的横向页由引擎按列数自行判定，故丢掉 rotating 的 sideways 方向不影响排版。)"""
    tex = re.sub(r'(?<![\\A-Za-z])[LRC]\{([^{}]*(?:cm|mm|pt|in|em|\\linewidth|\\textwidth|\\columnwidth)[^{}]*)\}',
                 r'p{\1}', tex)
    tex = re.sub(r'\\begin\{sidewaystable\}(\[[^\]]*\])?', r'\\begin{table}', tex)
    tex = tex.replace(r'\end{sidewaystable}', r'\end{table}')
    tex = _strip_resizebox(tex)
    tex = _expand_multicolumn(tex)
    tex = re.sub(r'\\(?:scriptsize|footnotesize)\s?', NOTE_SENTINEL, tex)  # 表脚注标记(渲染为 9pt)
    return tex

def _find_input(target, base_dir, resource_paths):
    cands = [target] if target.lower().endswith(('.tex','.ltx')) else [target, target+'.tex']
    for d in [base_dir] + list(resource_paths):
        for c in cands:
            p = c if os.path.isabs(c) else os.path.join(d, c)
            if os.path.isfile(p): return p
    bn = os.path.basename(cands[-1])                 # 回退：按 basename 在 resource-path 子树里找(容错作者写错的相对子目录)
    for d in resource_paths:
        for root,_,files in os.walk(d):
            if bn in files:
                hit = os.path.join(root, bn)
                sys.stderr.write(f'[subfmt] \\input{{{target}}} 未按原路径找到，改用 {hit}\n')
                return hit
    return None

def inline_inputs(tex, base_dir, resource_paths, depth=0):
    r"""递归内联 \input{file}，使 preprocess_latex 能作用到被包含的表文件
    (否则 sideways/resizebox 在 \input 的子文件里漏改、宽表全丢)。文献库(thebibliography)
    不内联，交给 pandoc / --bibliography，避免重复。"""
    if depth > 15: return tex
    def repl(m):
        p = _find_input(m.group(1).strip(), base_dir, resource_paths)
        if not p: return m.group(0)
        content = open(p, encoding='utf-8').read()
        if r'\begin{thebibliography}' in content: return m.group(0)
        return inline_inputs(content, os.path.dirname(p), resource_paths, depth+1)
    return re.sub(r'\\input\s*\{([^}]*)\}', repl, tex)

def run_pandoc_json(path, resource_paths):
    tmp = None; cmd_path = path
    if path.lower().endswith('.tex'):
        base = os.path.dirname(os.path.abspath(path))
        tex = inline_inputs(open(path, encoding='utf-8').read(), base, resource_paths)
        tex = preprocess_latex(tex)
        tmp = os.path.join(base, '._subfmt_pre.tex')  # 同目录保剩余 \input/图片基准
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
    ap.add_argument('--landscape-fit', type=float, default=LANDSCAPE_FIT,
                    help=f'自然总宽 > 纵向页宽×此值 → 转横向(默认 {LANDSCAPE_FIT})；调大=更倾向纵向，调小=更倾向横向')
    ap.add_argument('--landscape-mincols', type=int, default=99,
                    help='可选硬性 override：列数 >= N 一律横向(默认 99=关；默认只按 --landscape-fit 的内容宽度判定)')
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
    ctx = Ctx(rp, a.landscape_mincols, pkg, a.landscape_fit)

    # 参考文献须在渲染正文前解析：① 建 cite key→编号映射(供正文上标引用) ② 决定插入位置
    items = None
    if a.bibliography:
        _CITE_KEY2NUM.clear()
        if a.bibliography.lower().endswith(('.tex','.bbl')):
            items = parse_bibitems(a.bibliography)                          # [(key, text)]
            _CITE_KEY2NUM.update({k: i for i, (k, _) in enumerate(items, 1)})
        else:
            refs = run_pandoc_json(a.bibliography, rp)
            items = [plain_text(b['c']) for b in refs['blocks'] if b.get('t') in ('Para','Plain')]

    body = [] if a.body_only else title_page(title, authors, affs, keywords, abstract)
    render_blocks(ast['blocks'], body, ctx, lead_break=not a.body_only)
    if items:
        # 参考文献排在首个 Supplementary 大节之前(真值顺序)；无补充材料则置于文末
        idx = ctx.supp_idx if ctx.supp_idx is not None else len(body)
        body[idx:idx] = references_block(items)

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
    document = document.replace(NOTE_SENTINEL, '')   # 兜底：清除任何残留脚注哨兵
    open(f'{pkg}/word/document.xml','w',encoding='utf-8').write(document)
    repack(pkg, a.output)
    print(f"[ok] {a.output}  (authors={len(authors)}, figs={ctx.fig_n+ctx.s_fig_n}, tables={ctx.tab_n+ctx.s_tab_n})")

if __name__ == '__main__':
    main()
