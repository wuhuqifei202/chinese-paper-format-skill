# chinese-paper-format-skill

> 将中文论文 .docx 文件按学术排版规范自动格式化 — 题目、四级标题、正文、脚注全覆盖。

## Purpose

这是一个中文学术论文自动排版工具。读取 .docx 论文文件，通过内容模式自动识别论文题目、一~四级标题和正文段落，然后按照中文学术排版规范逐一设置字体、字号、加粗、对齐、行距和缩进。同时格式化页下脚注（宋体小五号、每页重新编号）。

## Activation

你说任何跟「修改论文格式」相关的话都会激活此技能，不需要记命令：

**自然语言触发:**
- "使用修改论文格式skill，帮我把这篇论文排版成学术规范"
- "用论文格式化技能处理这个 docx 文件"
- "帮我把论文格式统一成中文规范"
- "修改论文格式：标题黑体四号居中，正文宋体五号"
- "这篇论文的排版需要规范化，帮我调整一下"
- "格式化这个 docx 论文"
- "批量调整这个文件夹里所有论文的排版"

只要请求中包含「论文」+「格式/排版/格式化/规范」或「修改/调整/统一」+「论文格式」就会激活。

## Usage

1. 先用 `--check` 检查文档结构
2. 确认标题检测正确后执行格式化
3. 正文首行缩进默认 2 字符，可用 `--body-indent 0` 关闭

```bash
# 检查结构
python scripts/format_paper.py 论文.docx --check

# 格式化
python scripts/format_paper.py 论文.docx -o 论文_已格式化.docx --body-indent 2

# markdown 中转工作流 (docx → md → 重建 docx, md 与 docx 双交付; 支持 .doc)
python scripts/run_pipeline.py --input 论文.docx --via-markdown --fix-citations

# 自定义格式规则 (按需换格式; 只写要改的元素, 未写的用默认)
python scripts/format_paper.py 论文.docx --rules 我的格式.json -o 输出.docx
python scripts/run_pipeline.py --input 论文.docx --via-markdown --rules 我的格式.json
```

**自定义格式规则**: 用户说"按 XX 要求调整格式"(如题目黑体二号加粗、正文小四
20磅行距) 时, 先看 `assets/rules.example.json` 示例, 为需求创建 JSON 配置,
用 `--rules` 传入。支持元素: title/author/abstract/keywords/h1–h4/body/footnote;
字段: font/size/bold/align/indent/line_spacing。细节见 `references/custom-rules.md`。
四级标题 (（1）) 自动检测为 h4, md 层用 `#####` 前缀。

**自定义注释体例**: 默认依据《法学引注手册》(2019)；用户未提出要求时按默认
修改。**用户无需手写 JSON** — 在对话框用文字描述即可, skill 把文字翻译成
`--citation-rules 体例.json` 或 `--preset 预设`。预设: `法学引注手册`(默认)、
`《法学家》《中外法学》注释体例`、`《中国法学》`(与法学引注手册一致)、
`《法商研究》`(与法学引注手册一致)、`《法学研究》`(法学引注手册 + 14 条要求)。
配置支持 `rules`(修复规则开关)、`conventions`(命名约定)、
`categories`(各类体例模板/示例)、`notes`(体例要求)。细节见
`references/citation-rules.md`; 用 `python scripts/citation_formatter.py
--show-style` 查看默认体例、`--show-style --preset 《法学研究》` 查看含 14 条
要求的体例。

**Markdown 中转工作流** (用户需要审查/修改内容, 或输入为 .doc 时优先使用):
输入先转换为结构化 markdown (段落/标题层级/脚注 `[^n]`), 可在 md 层修改
(引注修复、结构调整), 再由 md2docx 按规则表重建 .docx。md 不承载排版与
表格/图片, 这些在转换中丢失并在重建时统一套用。.doc 需 LibreOffice
(soffice), 未安装时提示安装。

## Implementation

Full skill definition, scripts, and references are in [SKILL.md](./SKILL.md). See SKILL.md for complete formatting rules, trigger keywords, and error handling details.

**版本与同步 (唯一事实源)**: 本目录 (~/.claude) 是 git 主仓库,
`~/.workbuddy/skills/chinese-paper-format-skill` 是单向同步的备用目录 (无 git, 禁止直接改)。
每次 `git commit` 由 post-commit hook 自动调用 `sync_to_workbuddy.py --apply --quiet`
同步; 也可手动 `python sync_to_workbuddy.py` (dry-run) / `--apply` (执行)。

## Files

- `SKILL.md` — Full skill definition (agentskills.io format)
- `scripts/format_paper.py` — Main formatting script with CLI (排版+引注)
- `scripts/citation_formatter.py` — Citation format checker/fixer (法学引注手册 2019)
- `scripts/citation_rules.py` — Citation style table (默认法学引注手册 + 法学家/中外法学/中国法学/法商研究/法学研究等期刊预设 + JSON config load/validate)
- `scripts/docx2md.py` — Conversion layer: docx → structured markdown
- `scripts/md2docx.py` — Conversion layer: markdown → rebuilt docx (rules-table formatting)
- `scripts/run_pipeline.py` — Single-command orchestrator (incl. --via-markdown + .doc via LibreOffice)
- `scripts/rules.py` — Custom rules table (JSON config load/validate, CN size table)
- `references/formatting-rules.md` — Detailed formatting rule reference
- `references/custom-rules.md` — Custom rules reference (elements/fields/example)
- `references/citation-rules.md` — Custom citation style reference (注释体例/rules/conventions)
- `assets/rules.example.json` — Example custom rules config
- `assets/citation_rules.example.json` — Example custom citation style config
- `requirements.txt` — Python dependencies
- `README.md` — Multi-platform installation
- `install.sh` — Cross-platform installer
- `sync_to_workbuddy.py` — 单向同步到 ~/.workbuddy 备用目录 (post-commit hook 自动调用; 不随同步复制)
