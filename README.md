# chinese-paper-format-skill

自动将中文论文 .docx 文件按照学术排版规范格式化。

## 格式化内容

- 论文题目: 黑体, 四号, 居中
- 一级标题 (一、): 宋体, 小四, 加粗, 居中
- 二级标题 ((一)): 楷体, 小四, 加粗, 缩进
- 三级标题 (1. ): 宋体, 五号, 加粗, 缩进
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
```

## Troubleshooting

- **"不支持的文件格式"**: 仅支持 .docx，.doc 文件请先用 Word 另存为 .docx
- **标题未被检测**: 检查原标题是否符合 "一、" / "（一）" / "1. " 格式
- **脚注未格式化**: 确认原文中的注释是通过 Word 脚注功能插入的，而非手动输入的文字
