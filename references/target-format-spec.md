# MeVO 投稿 Word 格式 — 逆向工程权威规格

> 真值文件：`mevo_manuscript.docx`（解包后 `word/document.xml`、`word/styles.xml`）
> 由多 agent 调研逆向，所有数值已对 XML 复核一致（pStyle 分布、9 处分页符、11 表 / 10 图 / 0 anchor、6 sectPr / 2 landscape、0 ind、sz/jc 分布、docGrid linePitch=312、0 pageBreakBefore 等）。
> 用途：submission-format 引擎实现的格式标准——要改格式改这里、引擎按此渲染。

---

## 第 1 部分 — 格式规格

### 1.0 总机制：clean 格式 = 全 Normal + 直接格式（最关键）

"扁平化"稿，**几乎不用语义命名样式承载格式**，全部格式（jc / sz / b / spacing）直接写在每个段落的 `pPr` 和 run 的 `rPr` 上。

| 指标 | 实测值 |
|---|---|
| 段落总数 | 1058 段 |
| `pStyle=Normal` | **1045 段（98.8%）** |
| `pStyle=address` | 10 段（单位 / 关键词 / 通讯邮箱 / Research-in-context 区） |
| `pStyle=author` | 1 段（作者行） |
| `pStyle=Heading*` | **0 处**（Heading1-9 仅在 styles.xml 定义、正文从不引用） |
| `pageBreakBefore` | **0 处**（分页一律用显式 `<w:br w:type="page"/>`） |
| `<w:ind>` / `w:hanging` | **各 0 处**（全文无任何缩进） |

**Normal 样式基线**（styles.xml:398-416）：
- `pPr`：`widowControl=false`；`spacing lineRule="auto" line="278" before="0" after="160"`（行距 ≈1.16x，段后 8pt，段前 0）；`jc="start"`（**默认左对齐**）
- `rPr`：主题字体等线（**但正文 run 普遍覆写为 Times New Roman + eastAsia 宋体**）；`sz="22"`（11pt）`szCs="24"`；`kern=2`

**字号分布**（半磅制 ÷2=pt）：`sz=18`（9pt）×1859 → 表格单元格+题注；`sz=22`（11pt）×57；`sz=28`（14pt）×39 → 大节标题+主标题。
**对齐分布**：`jc=both` ×185（正文段+标题）；`jc=center` ×835（多为表格单元格）。
**docGrid（行距视觉根因）**：每个 sectPr 带 `<w:docGrid type="lines" linePitch="312">`，把正文吸附到 15.6pt 行网格，渲染出 ≈1.5 行观感（即便 Normal 写的是 1.16x）。
**页面几何**：A4 纵向 `pgSz 11906×16838`；页边距 `left/right=1800`（1.25″）、`top/bottom=1440`（1″）。

### 1.1 Title page 标题页（document.xml 行 3–428）

**① 主标题**：14pt（sz=28/szCs=32）粗体（b/bCs）；**左对齐 start**（无本地 jc，继承 Normal）；`pStyle=Normal`；TNR + 宋体。
**② 作者行**（`pStyle=author`）：单段含全部作者 "First Author^a, Second Author^a, …, Corresp Author A^a,*, Corresp Author B^a,**"（示例占位）。**斜体**（作者名继承 author 样式的 `<w:i/>`，渲染为斜体——经真值 PDF 实测确认，run 不取消斜体；此前规格误记为"不斜体"，已更正）；单位字母 `a` 为 `vertAlign=superscript`；通讯 `*`/`**` 上标；11pt（本地 sz=22 覆写 author 默认 12pt）；`jc=both`；`spacing before=0 after=156`。
**③ 单位行**（`pStyle=address`）：**自动编号** `numPr numId=1 ilvl=0` → `numFmt=lowerLetter`、`lvlText="%1."` 即 **"a."**；`ind start=360 hanging=360`；11pt；`jc=both`。
**④ 通讯作者邮箱行 ×2**（`pStyle=address`，`jc=both`，`spacing before=156 after=0`，11pt）。
**⑤ 收尾分页符**：空 Normal 段（rPr 带 sz=28 b）内含 `<w:br w:type="page"/>`（行 426）。

### 1.2 Abstract / Summary 区（行 429–634）

标题文字是 **"Summary"**（非 "Abstract"）。

| 元素 | 字号 | 加粗 | 对齐 | 段距/行距 | pStyle |
|---|---|---|---|---|---|
| "Summary" 大标题 | 14pt | 是 | jc=both | before=0 after=160, line=278 | Normal |
| 4 小标题 Background/Methods/Findings/Interpretation | 11pt（继承） | 是 | jc=both | line=278，无 spacing 覆盖 | Normal |
| 摘要正文段 | 11pt | 否 | jc=both | before=0 after=160 | Normal |
| Keywords 行 | 11pt | "Keywords:" 标签加粗 | jc=both | before=156 after=0, line=360(1.5x) | address |

- 4 小标题**各自独占一段**、正文另起一段（不是 run-in 内联）。
- Keywords：" Medium-vessel occlusion stroke; Endovascular thrombectomy; Prognosis; Functional topology; CT perfusion"（**分号分隔**）。
- "Summary" 紧接第 1 个分页符独占新页。

### 1.3 正文段落与标题机制

- **正文段**：`pStyle=Normal` + `jc=both` + run 直接 TNR（宋体）；无 spacing/sz 覆盖 → 1.16x、11pt 全继承 Normal。
- **大节标题**：`pStyle=Normal` + 直接 `<w:b/>` + 直接 `sz=28`（14pt，`szCs=32`） + `jc=both`；`after=160`、**`before=0`（真值全文 `before=480` 出现 0 次——此前规格的"后续大节 before≈480"系误记，已更正：大节靠分页符或前段 after=160 分隔，run-on 声明段标题亦 before=0）**；TNR。清单：主标题、Introduction / Methods / Results / Discussion / Conclusion / Contributors / Data sharing statement / Declaration of interests / Acknowledgements / References / Supplementary Material / Research in context / Summary。
- **小节标题**（如 "Study Design and Population"）：`pStyle=Normal` + 直接 `<w:b/>` + **无 sz 覆盖（11pt）** + `jc=both` + **无 spacing 覆盖** → 与正文同字号同行距，**仅靠加粗区分、不加段前距**。
- **Research in context 区**（例外用 address 样式）：主标题=一级标题格式；3 小节标题=`address`+b+11pt+jc=both+before=156 after=0；3 正文段=`address`+不粗+11pt+jc=both+before=156 after=0+行距 line=360(1.5x)。
- **四声明段**（Contributors / Data sharing / Declaration of interests / Acknowledgements）：标题=一级标题格式；正文=Normal+11pt+jc=both+TNR；**四段连排同一页**（彼此无分页符，紧跟 Conclusion 后）。

### 1.4 九个分页符位置（恰好 9 个 `<w:br w:type="page"/>`）

均嵌在"占位空段"内（段 pPr 带标题级 rPr：b+sz=28+szCs=32；段内第一个 run 空、第二个 run 仅含分页符）。

| # | 行号 | 把谁推到新页 |
|---|---|---|
| 1 | 426 | Summary 标题 |
| 2 | 657 | Research in context |
| 3 | 839 | Introduction |
| 4 | 1056 | Methods |
| 5 | 1731 | Results |
| 6 | 18847 | Discussion |
| 7 | 19112 | Conclusion |
| 8 | 19346 | References（四声明段连排在 Conclusion 后同页；References 才另起页） |
| 9 | 19811 | Supplementary Material |

### 1.5 参考文献（References）

- 标题：14pt 加粗（一级标题格式）。
- 列表：**30 条**（编号 1–30），每条独立 `<w:p>`，`pStyle=Normal`、TNR、**11pt**、line=278、before=0 after=160。
- **编号为手打字面文本** "1. ""2. "…（**非自动编号、无 numId**）。
- **无悬挂缩进、无任何缩进**。
- **对齐 = 左对齐 start**（参考段 pPr 无 jc，回退 Normal 的 jc=start）—— **与正文段显式 jc=both 不同，易错点**。
- 格式："1. 作者. 标题. 期刊 年; 卷: 页. doi"。

### 1.6 图与表格式

**数量**：11 表（Table 1–4 + S1–S7）；10 图（Figure 1–5 + S1–S5），**全部 `<wp:inline>` 内嵌、0 浮动**。

**表格（全框线网格表，非三线表）**：
- 每单元格 **四边 `single sz=4 color=000000`**；单元格统一 **9pt**（sz=18）；首列 jc=both（左）、数据列 jc=center；cell vAlign=center。
- **单元格段落 spacing `lineRule=auto line=240 before=0 after=0`**（单倍行距、无段距——不可继承 Normal 的 278/160，否则行高过松；引擎在 `cell_xml` 显式写出）。
- 表块 `tblPr jc=center` + `tblLayout=fixed` + `tblCellMar top/bottom=0 start/end=108`。
- 表宽：多数 `tblW 5000 pct`（满宽）；Table 3 用 `13921 dxa`（配横向页）。
- **表头加粗是通则**（Table 2/3/4 及全部 S 表），**唯独 Table 1 表头不加粗（唯一例外）**。

**题注**：
- **表标题在表【上方】**，"Table N. " 标签加粗（b+sz=18）+ 描述非粗（sz=18），jc=both，spacing before=120 after=160；行距 **line=240（单倍，按用户偏好）**——真值原为 line=360(1.5x)，引擎已按用户要求改 1.0。
- **图标题在图【下方】**，"Figure N. " 标签加粗（b+sz=18）+ 描述非粗，jc=both，spacing line=240 before=0 after=120；图所在段 jc=center。
- **表脚注在表下方**，9pt，带 iCs，jc=both，before=0 after=120。

**横向页**：宽表 Table 2/3/4 放横向 A4（`sectPr orient=landscape 16838×11906`，margin left/right=1080）；其余与补充材料为纵向 A4。全文 6 个 sectPr（2 横 + 4 纵），每个带 docGrid linePitch=312。
（注：横向页由引擎按**内容自然宽**判定（见 §1.7 的 `--landscape-fit`，默认 1.6），故 landscape 节数随实际宽表数量变化，与真值的"2 横"不必相等——真值是更早的数据版本、宽表更少。）

### 1.7 正文引用 / 补充材料编号 / 参考文献位置（引擎行为，2026-06 校正补充）

- **正文引用 = 上标编号**：`\cite{key}` 渲染为上标 run（仅 `vertAlign=superscript` + TNR、**无 sz 覆盖**，继承 11pt）。编号 = 该 key 在 `--bibliography`（`\bibitem` 顺序）中的 1-based 序号。多 key 排序后**连续段用连字符**（`2-3`、`5-7`）、**跳号用逗号无空格**（`7,9`），可组合（`5-6,24-25`）。与标题页的单位字母上标（带 `sz=22`）区分。
- **补充材料 S 编号**：进入首个以 "Supplementary" 开头的大节后，图/表切到 **`Table S1…`/`Figure S1…`**（独立 S 计数器），并剥除 caption 自带的 `Table N./Table S1.` 前缀（消双前缀）。
- **参考文献位置**：`--bibliography` 生成的 References 块（分页符 + 14pt 标题 + 左对齐 `jc=start` 编号条目）插在**首个 Supplementary 大节之前**（真值顺序：…Acknowledgements → References → Supplementary…）；无补充材料时置于文末。
- **LaTeX 预处理（让 pandoc 能解析）**：`\input` 先**递归内联**（否则被包含文件里的 sideways/resizebox 漏改、宽表全丢；文献库 thebibliography 不内联）→ `sidewaystable→table` + 去 `\resizebox{}{}{}` 包裹 + 展开 `\multicolumn`（否则整行分类小标题被 pandoc 丢弃）+ `L/R/C{宽}→p{宽}`。`\input` 找不到原路径时按 basename 在 `--resource-path` 子树里兜底搜索。
- **表脚注识别**：`\scriptsize`/`\footnotesize` 预处理期换成哨兵 → 段首带哨兵的段落判为表脚注，渲染为 9pt iCs（见 1.6）。
- **横/纵向按内容宽度判定**（`_is_wide`，取代旧的"列数 ≥ N"）：各列自然宽 = 最长单元格字符数 × `CHAR_DXA(95)` + 内边距；**自然总宽 > 纵向可用宽(8306) × `--landscape-fit`(默认 1.6) 才转横向**，否则压进纵向页。故"列多但纵向放得下"的表（如 7 列的 Table S6/S7，自然宽≈1.25–1.29×纵向）保持纵向；只有真宽的 Table 2/3/4（1.73–2.80×）转横向。`--landscape-mincols`（默认 99=关）为可选硬性 override。
- **列宽按内容自动分配**（`render_table`/`_col_widths`）：列宽 ∝ 该列最长单元格字符数（自然宽）；自然宽合计 ≤ 可用宽则等比放大（全列 ≥ 自然宽、不换行），否则等比压缩，但**每列不低于其最长单词宽**（数字/单词不折断，如 "73" 不被拆成 7/3）。目的：尽量不换行、表整体高度最短。`tblW`/`gridCol`/`tcW` 均用 dxa（纵向 8306 / 横向 13958），不再等分。
- **横向表脚注归位 + 无空白页**（`render_blocks` 用 `ctx.landscape_close`）：宽表的 `LANDSCAPE` 收尾分节段**推迟到表脚注之后**插入（脚注/空段并入横向节、遇正文内容才收尾），状态放 `ctx` 上以跨 pandoc 的 `Div[Table]` 浮动包裹递归共享——否则表脚注被挤到表后的纵向页。**相邻横向宽表**（`_is_wide_block` 穿透 Div 判定）续在**同一横向节**、不另插 `pre` 纵向空段（否则两表间夹出空白纵向页）；**横向表后紧跟的大节标题**省掉 `pagebreak()`（横向分节符已换页，否则多一空白页）。RIC 正文/小标题须显式 `sz=22`（address 样式默认 12pt，不写就继承成 12pt）。
