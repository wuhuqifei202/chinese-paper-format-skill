# chinese-paper-format-skill

> 将中文论文 .docx 文件按学术排版规范自动格式化 — 题目、三级标题、正文、脚注全覆盖。

## Purpose

这是一个中文学术论文自动排版工具。读取 .docx 论文文件，通过内容模式自动识别论文题目、一/二/三级标题和正文段落，然后按照中文学术排版规范逐一设置字体、字号、加粗、对齐、行距和缩进。同时格式化页下脚注（宋体小五号、每页重新编号）。

## Activation

Invoke with `/chinese-paper-format-skill` on platforms that support slash commands, or describe the task naturally.

**Example queries:**
- "帮我把这篇论文的格式调整成学术规范"
- "格式化这个 docx 论文，标题用黑体四号居中"
- "批量调整这个文件夹里所有论文的排版"
- "检查一下这篇论文的格式是否符合中文论文规范"

## Usage

1. 先用 `--check` 检查文档结构
2. 确认标题检测正确后执行格式化
3. 推荐加 `--body-indent 2` 设置正文首行缩进

```bash
# 检查结构
python scripts/format_paper.py 论文.docx --check

# 格式化
python scripts/format_paper.py 论文.docx -o 论文_已格式化.docx --body-indent 2
```

## Implementation

Full skill definition, scripts, and references are in [SKILL.md](./SKILL.md). See SKILL.md for complete formatting rules, trigger keywords, and error handling details.

## Files

- `SKILL.md` — Full skill definition (agentskills.io format)
- `scripts/format_paper.py` — Main formatting script with CLI
- `scripts/run_pipeline.py` — Single-command orchestrator
- `references/formatting-rules.md` — Detailed formatting rule reference
- `requirements.txt` — Python dependencies
- `README.md` — Multi-platform installation
- `install.sh` — Cross-platform installer
