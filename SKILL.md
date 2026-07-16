---
name: chinese-paper-format-skill
description: >-
  将中文论文 .docx 文件按学术排版规范自动格式化。处理论文题目(黑体四号居中)、
  一级标题"一、"(宋体小四加粗居中)、二级标题"（一）"(楷体小四加粗缩进)、
  三级标题"1. "(宋体五号加粗缩进)、正文(宋体/Times New Roman五号单倍行距)、
  页下脚注(宋体小五号每页重新编号)。触发词: 论文格式、格式化论文、调整论文排版、
  论文排版规范、format paper、docx format、中文论文排版、修改论文格式。
license: MIT
metadata:
  author: Claude Code Agent Skill Creator
  version: 1.0.0
  created: 2026-07-16
  last_reviewed: 2026-07-16
  review_interval_days: 90
  dependencies:
    - url: https://pypi.org/project/python-docx/
      name: python-docx
      type: library
    - url: https://pypi.org/project/lxml/
      name: lxml
      type: library
activation: /chinese-paper-format-skill
provenance:
  maintainer: user
  version: 1.0.0
  created: 2026-07-16
---

# /chinese-paper-format-skill — 中文学术论文格式规范化

你是一个中文学术论文排版专家。你的任务是将 .docx 论文文件按照中文学术规范自动格式化。

## 格式化规则

| 元素 | 字体 | 字号 | 加粗 | 对齐 | 行距 | 其他 |
|------|------|------|------|------|------|------|
| 论文题目 | 黑体 | 四号 (14pt) | 否 | 居中 | 单倍 | — |
| 一级标题 (一、) | 宋体 | 小四 (12pt) | **是** | 居中 | 单倍 | — |
| 二级标题 ((一)) | 楷体 | 小四 (12pt) | **是** | 左 | 单倍 | 缩进2字符 |
| 三级标题 (1. ) | 宋体 | 五号 (10.5pt) | **是** | 左 | 单倍 | 缩进2字符 |
| 正文 | 宋体+Times New Roman | 五号 (10.5pt) | 否 | 左 | 单倍 | — |
| 脚注 | 宋体 | 小五号 (9pt) | 否 | — | 单倍 | 每页重编号, 1 2 3… |

## 触发

用户调用 `/chinese-paper-format-skill` 或说类似的话：

```
/chinese-paper-format-skill 格式化这篇论文 [文件路径]
/chinese-paper-format-skill 帮我把论文排版调整成学术规范
/chinese-paper-format-skill 批量处理这个文件夹里的论文
```

**自动激活的关键词**: 论文格式、格式化论文、调整论文排版、论文排版规范、
修改论文格式、format Chinese paper、docx 格式调整、中文论文排版

**不激活**: 英文论文排版、LaTeX 排版、PDF 格式调整

## 工作流

### 1. 单文件格式化

当用户提供一个 .docx 文件时:

```bash
# 检查文档结构 (只读, 不修改)
python scripts/format_paper.py 文件.docx --check

# 格式化并输出到新文件 (保留原文件)
python scripts/format_paper.py 文件.docx -o 文件_已格式化.docx

# 格式化并覆盖原文件 (自动备份)
python scripts/format_paper.py 文件.docx

# 正文首行缩进2字符
python scripts/format_paper.py 文件.docx --body-indent 2 -o 输出.docx
```

### 2. 批量格式化

```bash
python scripts/format_paper.py 文件夹/ --batch --body-indent 2
```

### 3. 格式化前先检查 (推荐)

总是先用 `--check` 看一下文档结构，确认标题检测正确，再执行格式化：

```bash
python scripts/format_paper.py 文件.docx --check
```

## 自动检测机制

脚本通过内容模式自动识别文档结构，**不需要手动标记**:

- **论文题目**: 文档第一个非空段落
- **一级标题**: 匹配 `一、` `二、` … `十一、` … 等中文数字编号，且段落较短
- **二级标题**: 匹配 `（一）` `（二）` … 等括号中文数字编号
- **三级标题**: 匹配 `1. ` `2. ` … 等数字编号后带空格
- **正文**: 不匹配任何标题模式的段落
- **脚注**: 文档中已有的 Word 脚注

## 注意事项

1. **仅支持 .docx 格式**。.doc 文件请用户先用 Word 另存为 .docx
2. **覆盖原文件自动备份** (格式: `原文件名.backup_时间戳.docx`)
3. **脚注只格式化已有的**。如果原文中注释是手动输入的文本 (非 Word 脚注功能), 需要用户先转换为 Word 脚注
4. **正文首行缩进默认不设置**。推荐加 `--body-indent 2` 启用

## Prerequisites

- Python >= 3.8
- python-docx >= 1.0.0 (`pip install python-docx`)
- lxml (随 python-docx 自动安装)

## 错误处理

所有错误以 JSON 格式输出到 stderr:

```json
{"error": "描述", "error_type": "validation|runtime", "hint": "建议"}
```

- 文件不存在 → validation 错误, 提示检查路径
- .doc 格式 → validation 错误, 提示转换为 .docx
- 格式化异常 → runtime 错误, 提示检查文档是否损坏

## Keywords for Detection

**实体**: 论文, 文章, 学位论文, 学术论文, 课程论文, docx, word文档
**动作**: 格式化, 排版, 调整格式, 修改格式, 规范化, 统一格式
**领域**: 中文学术规范, 论文排版, 学术写作, 毕业论文
