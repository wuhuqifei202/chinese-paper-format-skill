# chinese-paper-format-skill

自动将中文论文 .docx 文件按照学术排版规范格式化。支持 markdown 中转工作流
(doc/docx → 结构化 md → 重建 docx, md 与 docx 双交付) 与 .doc 旧格式
(LibreOffice 自动转换)。

## 格式化内容

- 论文题目: 黑体, 四号, 居中
- 一级标题 (一、): 宋体, 小四, 加粗, 居中
- 二级标题 ((一)): 楷体, 小四, 加粗, 缩进
- 三级标题 (1. ): 宋体, 五号, 加粗, 缩进
- 四级标题 ((1)): 宋体, 五号, 加粗, 缩进
- 正文: 宋体 + Times New Roman, 五号, 单倍行距
- 脚注: 宋体, 小五号, 每页重编号

## Installation

### Universal Path (works with 6+ tools)

```bash
git clone <repo-url> ~/.agents/skills/chinese-paper-format-skill
```

### Using install.sh (Recommended)

```bash
chmod +x install.sh
./install.sh                          # Auto-detect platform
./install.sh --platform claude-code   # Claude Code
./install.sh --platform workbuddy    # WorkBuddy (自动格式适配)
./install.sh --all                    # All detected platforms
./install.sh --dry-run                # Preview without installing
```

### Manual Installation

| Platform | Copy to |
|---|---|
| Universal | `~/.agents/skills/chinese-paper-format-skill/` |
| Claude Code | `~/.claude/skills/chinese-paper-format-skill/` |
| GitHub Copilot | `.github/skills/chinese-paper-format-skill/` |
| Cursor | `.cursor/rules/chinese-paper-format-skill/` |
| WorkBuddy | `~/.workbuddy/skills/chinese-paper-format-skill/` (格式适配版: SKILL.md 转换 + _skillhub_meta.json) |

## Prerequisites

- Python >= 3.8
- python-docx >= 1.0.0: `pip install python-docx`

## Usage

```bash
# Check document structure (read-only)
python scripts/format_paper.py paper.docx --check

# Format to new file
python scripts/format_paper.py paper.docx -o paper_formatted.docx --body-indent 2

# Format in-place (auto-backup)
python scripts/format_paper.py paper.docx --body-indent 2

# Batch process
python scripts/format_paper.py papers/ --batch --body-indent 2

# One-command pipeline
python scripts/run_pipeline.py --input paper.docx --output paper_formatted.docx

# Markdown via-workflow: docx → md → rebuilt docx (dual deliverable: .md + .docx)
python scripts/run_pipeline.py --input paper.docx --via-markdown --fix-citations

# .doc legacy format (requires LibreOffice)
python scripts/run_pipeline.py --input paper.doc --via-markdown

# Step-by-step conversion layer
python scripts/docx2md.py paper.docx -o paper.md     # docx → structured markdown
python scripts/md2docx.py paper.md -o paper.docx     # markdown → rebuilt docx

# Custom format rules (JSON config, see references/custom-rules.md)
python scripts/format_paper.py paper.docx --rules my-rules.json -o paper_formatted.docx
python scripts/run_pipeline.py --input paper.docx --via-markdown --rules my-rules.json
```

Custom rules let you convert to your own format requirements instead of the
preset rules: write only the elements you want to change in a JSON config
(e.g. `{"title": {"size": "二号", "bold": true}, "body": {"size": "小四",
"line_spacing": "20磅"}}`); unspecified elements/fields fall back to defaults.
`_`-prefixed keys are comments. See [references/custom-rules.md](references/custom-rules.md)
and the example file [assets/rules.example.json](assets/rules.example.json).

Custom citation style (注释体例): by default citations are normalized per
《法学引注手册》(2019). Users describe their citation style in plain text in
the chat — no hand-written JSON needed; the skill translates it into
`--citation-rules 体例.json` or `--preset` (法学引注手册 / 《法学家》《中外法学》注释体例 /
《中国法学》 / 《法商研究》 / 《法学研究》).
Config only needs the parts you change (`rules` toggles, `conventions`,
`categories`, `notes`). See [references/citation-rules.md](references/citation-rules.md)
and [assets/citation_rules.example.json](assets/citation_rules.example.json).

```bash
# 按自定义注释体例修复引注
python scripts/format_paper.py paper.docx --fix-citations --citation-rules 体例.json -o out.docx
# 按《法学家》《中外法学》注释体例修复引注
python scripts/format_paper.py paper.docx --fix-citations --preset 《法学家》《中外法学》注释体例 -o out.docx
# 按《法学研究》注释体例修复引注（法学引注手册 + 14 条要求）
python scripts/format_paper.py paper.docx --fix-citations --preset 《法学研究》 -o out.docx
# 查看默认注释体例 (法学引注手册)
python scripts/citation_formatter.py --show-style
```

The markdown via-workflow converts the input to structured markdown first
(paragraphs, heading levels, footnotes as `[^n]` definitions), allows editing
the content, then rebuilds the .docx with formatting applied from the rules
table. Formatting (font/size/alignment) is NOT preserved through markdown —
it is re-applied on rebuild. Tables/images/comments are not supported in
markdown mode.

## Development & Sync (唯一事实源)

**主仓库**: `~/.workbuddy/skills/chinese-paper-format-skill` (git 仓库, 唯一事实源)。
**部署目录**: `~/.claude/skills/chinese-paper-format-skill` — 由主仓库单向同步,
**禁止直接在部署目录修改或提交** (该目录无 git 仓库, 曾因双仓分叉导致功能落后)。

同步机制:
- `python sync_to_claude.py` — dry-run 预览同步计划
- `python sync_to_claude.py --apply` — 执行同步 (复制/更新/删除目标中多余文件,
  排除 `.git*` / `__pycache__` / `.pytest_cache` / `*.old` / `*.pyc` / 本脚本自身)
- **post-commit hook**: 每次 `git commit` 后自动同步, 无需手动操作

开发流程: 改主仓库 → `git commit` → (hook 自动同步) → 部署目录即更新。
新增/改名文件后若目标中出现被误删除的部署文件, 以主仓库 git 历史为准恢复。

## Troubleshooting

- **".doc 无法转换"**: --via-markdown 模式下 .doc 需要安装
  [LibreOffice](https://www.libreoffice.org/) (soffice)，未安装时会提示；也可先用
  Word 另存为 .docx
- **标题未被检测**: 检查原标题是否符合 "一、" / "（一）" / "1. " 格式
- **脚注未格式化**: 确认原文中的注释是通过 Word 脚注功能插入的，而非手动输入的文字
