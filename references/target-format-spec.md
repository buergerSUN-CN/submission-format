# MeVO 投稿 Word 格式 — 逆向工程权威规格

> 真值文件：`mevo_manuscript.docx`（解包后 `word/document.xml`、`word/styles.xml`）
> 由多 agent 调研逆向，所有数值已对 XML 复核一致（pStyle 分布、9 处分页符、11 表 / 10 图 / 0 anchor、6 sectPr / 2 landscape、0 ind、sz/jc 分布、docGrid linePitch=312、0 pageBreakBefore 等）。
> 用途：重写 submission-format skill 的格式标准。

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
- **表标题在表【上方】**，"Table N. " 标签加粗（b+sz=18）+ 描述非粗（sz=18），jc=both，spacing line=360 before=120 after=160。
- **图标题在图【下方】**，"Figure N. " 标签加粗（b+sz=18）+ 描述非粗，jc=both，spacing line=240 before=0 after=120；图所在段 jc=center。
- **表脚注在表下方**，9pt，带 iCs，jc=both，before=0 after=120。

**横向页**：宽表 Table 2/3/4 放横向 A4（`sectPr orient=landscape 16838×11906`，margin left/right=1080）；其余与补充材料为纵向 A4。全文 6 个 sectPr（2 横 + 4 纵），每个带 docGrid linePitch=312。
（注：横向页由引擎按"列数 ≥ `--landscape-mincols`（默认 6）"判定，故 landscape 节数随实际宽表数量变化，与真值的"2 横"不必相等——真值是更早的数据版本、宽表更少。）

### 1.7 正文引用 / 补充材料编号 / 参考文献位置（引擎行为，2026-06 校正补充）

- **正文引用 = 上标编号**：`\cite{key}` 渲染为上标 run（仅 `vertAlign=superscript` + TNR、**无 sz 覆盖**，继承 11pt）。编号 = 该 key 在 `--bibliography`（`\bibitem` 顺序）中的 1-based 序号。多 key 排序后**连续段用连字符**（`2-3`、`5-7`）、**跳号用逗号无空格**（`7,9`），可组合（`5-6,24-25`）。与标题页的单位字母上标（带 `sz=22`）区分。
- **补充材料 S 编号**：进入首个以 "Supplementary" 开头的大节后，图/表切到 **`Table S1…`/`Figure S1…`**（独立 S 计数器），并剥除 caption 自带的 `Table N./Table S1.` 前缀（消双前缀）。
- **参考文献位置**：`--bibliography` 生成的 References 块（分页符 + 14pt 标题 + 左对齐 `jc=start` 编号条目）插在**首个 Supplementary 大节之前**（真值顺序：…Acknowledgements → References → Supplementary…）；无补充材料时置于文末。
- **LaTeX 预处理（让 pandoc 能解析）**：`\input` 先**递归内联**（否则被包含文件里的 sideways/resizebox 漏改、宽表全丢；文献库 thebibliography 不内联）→ `sidewaystable→table` + 去 `\resizebox{}{}{}` 包裹 + 展开 `\multicolumn`（否则整行分类小标题被 pandoc 丢弃）+ `L/R/C{宽}→p{宽}`。`\input` 找不到原路径时按 basename 在 `--resource-path` 子树里兜底搜索。
- **表脚注识别**：`\scriptsize`/`\footnotesize` 预处理期换成哨兵 → 段首带哨兵的段落判为表脚注，渲染为 9pt iCs（见 1.6）。

---

## 第 2 部分 — 当前 pandoc 输出的差距清单

> "mine" = 旧的 pandoc + reference.docx 输出（已废弃路线）。

### 2.1 格式 / 机制差距（按严重度）

| 严重度 | 维度 | 目标真值 | pandoc 输出 |
|---|---|---|---|
| HIGH | 段落样式机制 | 1045 Normal + 直接格式 | 12 种语义样式（Compact/BodyText/Heading1/2/Abstract…）机制相反 |
| HIGH | 分页符 | 9 个 | 0 个 |
| HIGH | 分节/横向页 | 6 sectPr（含 2 landscape 放宽表） | 1 纵向，无横向，宽表溢出 |
| HIGH | docGrid 行网格 | 每节 linePitch=312（≈1.5x 观感） | 无，正文 1.16x 明显更紧 |
| HIGH | 标题页作者块 | 单段、上标字母 a、*/** 通讯、单位 address 自动 a. 编号 | 每位作者各占一居中 Author 段，无上标无标记 |
| HIGH | 摘要 | "Summary" 14pt 粗 + 独立粗体小标题 + 正文 11pt jc=both | "Abstract" 居中 10pt(sz=20)，小标题做成 run-in 内联 |
| HIGH | 图/表数量 | 10 图 / 11 表 | 5 图 / 2 真表（其余为空壳） |
| HIGH | References | 完整章节 + 30 条 | 完全缺失 |
| MEDIUM | 表格直接格式 | 满宽 + 四边框 + 9pt + 数据列居中 | tblStyle + auto 宽 + 仅 firstRow 底线 + 11pt 左 |
| MEDIUM | 图注/表注 | 9pt，"Figure N./Table N." 加粗前缀 | 继承 11pt，无编号前缀 |
| MEDIUM | 标题对齐/间距 | 主标题 jc=start(左)；大节 before=0/480 | Title 居中；Heading1 before=360 jc=start |
| LOW | 小节标题 | 11pt 内联粗体，不加段前距 | Heading2 sz=22 + before=160 |
| LOW | Keywords | "Keywords:" 标签加粗 + 分号分隔 | 无标签，逗号分隔 |

### 2.2 漏掉的具体清单

- **缺 9 张表**：Table 2/3/4（主）+ Table S2-S7（补充）。pandoc 仅含 Table 1 与 S1。
- **缺 5 张图（图像）**：Figure S1-S5（仅题注无图像）。主图 1-5 齐。
- **缺整个 References 章节**（30 条）。
- **缺 2 个横向 landscape 小节**。
- **缺全部 9 个分页符**。
- **缺标题页结构件**：上标机构字母、机构地址段、两行 Corresponding authors。
- **缺 "Keywords:" 标签**。

> ⚠️ 数据版本（内容，非格式）：逆向所用样式 docx 与后续测试稿是两个不同的数据版本（样本量与统计量不同）。skill 只负责**格式**，内容忠实于输入，不涉及具体数据。

---

## 第 3 部分 — skill 重做方案

**核心矛盾**：目标要 ① clean（全 Normal+直接格式）② 9 分页符 ③ 完整结构化 title page ④ 全部 10 图/11 表/横向节/References；而 pandoc+reference-doc 必然产语义样式且丢 title page/分页/部分图表——机制根本相反。

### 路线对比

- **路线 A — 以 `mevo_manuscript.docx` 为骨架模板、脚本灌入新内容**：clean/分页/title page/图表壳天然 100% 保真；但对任意新稿内容映射困难、多输入弱、强耦合具体文档。clean 5 / 分页 5 / title 5 / 图表 5(沿用)~1(换新) / 多输入 **1**。
- **路线 B — pandoc 后处理 XML（压平 pStyle→Normal+内联、插分页、重建 title page、重写表格…）**：保留 pandoc 多输入与正文/引用解析；但**最复杂最脆**，等于既付 pandoc 又付重建。clean 4 / 分页 4 / title 3 / 图表 3 / 多输入 **5**。
- **路线 C — python-docx 程序化逐段重建**：完全掌控、与真值逐位对齐；短板是需自建多输入解析。clean 5 / 分页 5 / title 5 / 图表 5(前提输入抽全) / 多输入 3。

### 推荐：路线 C 为主，借鉴 A 的"模板锚点"兜底

理由：① 目标本质是"格式规格"而非"某一篇"，A 泛化 1/5 只适合做 golden 对照；② B 的最大成本（压平+内联）正是最脆处，等于双重付费；③ C 在 clean/分页/title page 三个 HIGH 维度满分，正中差距最大处；短板"多输入解析"可分层：**输入→结构化中间表示**（用 pandoc/markitdown 仅抽内容，不依赖其样式输出）+ **中间表示→目标 docx**（python-docx/oxml 按本规格组装）解耦。

### 落地核对清单（路线 C 成功判据，逐项可验）

- pStyle 全 Normal（+必要时 author/address）；0 个 Heading*；标题=Normal+直接 b+sz。
- 大节标题 sz=28、jc=both、before=0 after=160（首）/before≈480；小节仅 b、11pt、不加段前距。
- 正文 Normal+jc=both+run 直接 TNR；不写 spacing/sz 覆盖。
- 9 个 `<w:br type=page>`（Summary/RIC/Intro/Methods/Results/Discussion/Conclusion/References/Supplementary 前各一）；四声明段连排同页。
- title page：作者合段+上标 a+*/**+address 自动编号 lvlText="%1."+两行通讯邮箱。
- 摘要 "Summary" 14pt 粗；4 小标题独立粗体段 11pt；正文 11pt jc=both；Keywords 标签加粗+分号。
- References 30 条手打字面编号、11pt、**jc=start（左，非 both）**、无缩进。
- 表格四边 single sz=4、单元格 9pt、首列 jc=both/数据列 center、表块 jc=center、满宽 5000pct；表头加粗（Table 1 例外）。
- 题注：表上/图下，"Table N./Figure N." 标签加粗 9pt+描述 9pt、jc=both。
- 2 个 landscape sectPr（w=16838，margin 1080）放宽表；每 sectPr 带 docGrid linePitch=312。
- 内容完整：10 图（含 S1-S5）+ 11 表（含 Table 2/3/4 + S2-S7）+ References。
