---
name: submission-format
description: "把任意稿件(Markdown/LaTeX/Word/HTML 等)渲染成用户的标准投稿 Word 格式——clean 排版(全 Normal 段落+直接格式，非 Word 语义样式)、Times New Roman 11pt、A4、两端对齐正文、14pt 粗大节标题/11pt 粗小节标题、结构化 Summary 摘要、按章节分页、结构化 title page(作者+单位上标+通讯)、全框线 9pt 表格(宽表自动转横向页)、Figure/Table 题注、左对齐编号参考文献。当用户想把文稿/草稿转换/重排成‘我的投稿格式’/‘投稿 Word 格式’/manuscript 模板，或让论文匹配其投稿样式时使用。引擎是 scripts/build_submission.py，按 references/target-format-spec.md 的规格生成。不要用于与该投稿样式无关的普通 Word 文档(那用 docx skill)。"
---

# Submission Format 引擎

> ⚠️ **渲染投稿 docx 一律走 `build_submission.py`。** 禁止用 docx-js / python-docx 手工
> 复刻本格式——手工必漏 `docGrid linePitch=312`、分页符机制(易误用 `pageBreakBefore`)、
> 全 `Normal` 直接格式等。唯一例外是引擎不支持的布局(如改前/改后高亮对照)；此时手写须逐项
> 对照 `references/target-format-spec.md` 自检：`docGrid=312` / 0 `pageBreakBefore` / 0 `Heading*` / 全 `Normal`。

把任意稿件渲染成用户的 clean 投稿 Word 格式。**这是个通用引擎**：内容统一从
pandoc AST 抽取，按固定的格式规格(`references/target-format-spec.md`)程序化生成
全 `Normal` 段落 + 直接格式的 docx。不针对任何具体稿件硬编码。

组成：
- `scripts/build_submission.py` — 引擎(纯 Python，依赖 pandoc)
- `assets/reference_styles.docx` — 样式资产包(Normal/author/address 样式 + numbering，无内容)
- `references/target-format-spec.md` — 目标格式的权威逆向规格(逐项数值)

## 用法

```bash
python3 scripts/build_submission.py INPUT -o OUT.docx \
    [--resource-path DIR]...      # 图片/资源目录(可多次)；同时供 LaTeX \input 查找
    [--bibliography FILE]         # 参考文献源(.tex/.bbl 的 thebibliography 最佳，或 pandoc 可读格式)；
                                  #   .tex/.bbl 还会驱动正文 \cite 的上标编号(按 \bibitem 顺序)
    [--styles STYLES.docx]        # 样式资产包(默认 assets/reference_styles.docx)
    [--landscape-mincols N]       # 列数 >= N 的表判为宽表→横向页(默认 6)
    [--body-only]                 # 片段模式：跳过 title page/Summary，只渲染正文(首个大节不插前导分页符)
```

渲染**完整 manuscript** 用默认模式(出 title page + Summary)；渲染**片段**(如单独某节、改前/改后对照、补充材料块)加 `--body-only`，否则会被硬塞一个空 title page + Summary。

INPUT 支持 pandoc 能读的格式(.md/.tex/.docx/.html/.rst/...)。PDF 请先转文本再喂。

典型(LaTeX 多文件项目)：
```bash
python3 scripts/build_submission.py paper.tex -o paper_submission.docx \
    --resource-path figures --resource-path figures/supplementary \
    --bibliography reference/refs.tex
```

## 关键：skill 只渲染"输入中实际存在的内容"

这是通用引擎，不会去补全输入缺失的东西。**内容完整性是输入侧的责任**：

- **图**：AST 里 `\includegraphics{F1}`/`![](F1)` 的路径，引擎在 `--resource-path` 列出的目录里
  按 `F1`/`F1.png`/`F1.jpg` 查找并嵌入；找不到则插占位文字。无扩展名也能找(自动补)。
- **表**：渲染 pandoc AST 里**能解析到的**表。标准 `tabular`/`longtable`/Markdown 表能解析；
  `xltabular`/某些浮动环境或外部 `\input` 不到的表 pandoc 拿不到——那是输入问题，需在输入侧
  改成可解析的表或确保 `\input` 路径正确(从项目根运行、用 `--resource-path`)。
- **参考文献**：用 `--bibliography` 单独喂最稳(引擎自解析 `\bibitem`、自动编号)；
  或确保 `\input{refs}` 在 pandoc 能找到的位置。
- **多文件 LaTeX**：pandoc 按**输入文件所在目录**解析 `\input`。若 .tex 的 `\input{tables/..}`
  相对项目根，就从项目根运行(或把 .tex 复制到根)。

→ 转换前先检查输入是否自洽：`pandoc INPUT -t json | ...` 看章节/表/图是否齐全；缺的先在输入侧补好。

## 格式规格(引擎已内置，改格式改这里)

完整数值见 `references/target-format-spec.md`。要点：
- 全段 `Normal` + 直接格式(jc/sz/b 写在 pPr/run)，**不用 Heading 等语义样式**。
- 正文 TNR 11pt、两端对齐、行距 line=278、段后 160、docGrid linePitch=312。
- 大节标题 14pt 粗(jc=both)；小节标题 11pt 粗(仅加粗、无段前距)。
- 9 处分页(大节前各一)；声明段(Acknowledgements/Contributions/Conflicts/Data availability 等)连排同页。
- title page：作者合段+上标单位字母+`*/**`通讯+单位 address 自动"a."编号+通讯邮箱行。
- Summary：4 小标题独立粗体段+正文；Keywords 标签加粗+分号分隔。
- 表：全框线 sz=4、单元格 9pt(段落 line=240/before0/after0 单倍行距)、首列左/数据列中、表块居中、满宽；表头加粗；中点小数(`\cdot`→·)；宽表(≥N 列)横向。
- 题注：表上(line=360)/图下(line=240)，"Table N./Figure N." 加粗 9pt 前缀；表脚注 9pt iCs。
- 正文引用：`\cite{key}` → 上标编号(按 `--bibliography` 的 `\bibitem` 顺序)，连续段连字符 `2-3`、跳号逗号 `7,9`。
- 补充材料：进入 "Supplementary*" 大节后图/表自动切 `Table S1…`/`Figure S1…`。
- References：左对齐(jc=start，与正文 both 不同)、无缩进、11pt、手打编号；**排在首个 Supplementary 大节之前**。

## 验证

渲染对照(LibreOffice 有缓存，改后先 `pkill -f soffice`)：
```bash
pkill -f soffice; python3 ~/.claude/skills/docx/scripts/office/soffice.py --headless --convert-to pdf OUT.docx --outdir /tmp
pdftoppm -jpeg -r 95 /tmp/OUT.pdf /tmp/p && # 读 /tmp/p-*.jpg 抽查 title page / 摘要 / 表 / 图 / References
```

## LaTeX 表格自动预处理(引擎在送 pandoc 前做)

`\input` 先递归内联(使下述改写能作用到被包含的表文件) → `sidewaystable→table`、去 `\resizebox{}{}{}`
包裹、展开 `\multicolumn`(否则整行分类小标题被 pandoc 丢弃)、`L/R/C{宽}→p{宽}`。`\input{相对路径}`
找不到时按 basename 在 `--resource-path` 子树兜底搜索(容错作者写错的子目录，会打印改用了哪个文件)。
仍不支持：`xltabular`(换 `longtable`)。

## 已知限制

- 完整 title page 需输入含作者元数据：`.tex` 走 elsarticle 适配器(`\author/\affiliation/\ead`)；
  其它格式取 pandoc meta(可能只有作者名，无单位/通讯)。
- 宽表横向按列数(`--landscape-mincols`)判断，故 landscape 节数随实际宽表数变化(不强行等于真值的"2 横")。
- 表格列宽按 pandoc colspec 等分(全 `l` 列时均分)，标签列可能换行——与真值的人工列宽略有差异，不影响内容。
- Table 1 表头在真值里是**不加粗**的唯一例外，引擎对所有表头统一加粗(不识别该特例)。
- 引擎覆写整套格式，因此输入里的直接排版会被丢弃(这是目的)。

## 依赖

`pandoc`(必需)；可选 LibreOffice+pdftoppm(仅用于渲染验证)。
