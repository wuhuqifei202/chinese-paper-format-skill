---
name: chinese-paper-format-skill
description: >-
  将中文论文 .docx 文件按学术排版规范自动格式化。处理论文题目(黑体四号居中)、
  一级标题"一、"(宋体小四加粗居中)、二级标题"（一）"(楷体小四加粗缩进)、
  三级标题"1. "(宋体五号加粗缩进)、正文(宋体/Times New Roman五号单倍行距)、
  页下脚注(宋体小五号每页重新编号)、引注格式规范化(依据《法学引注手册》2019)。
  当用户说"修改论文格式"、"使用修改论文格式skill"、
  "论文格式修改"、"用论文格式化技能"、"帮我调整论文排版"、"格式化这篇论文"、
  "把论文排版改成学术规范"、"论文排版规范化"、"统一论文格式"、"docx论文格式调整"
  或任何与论文格式修改相关的请求时激活此技能。
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
| 摘要 | 楷体 | 小四 (12pt) | 否 | 左 | 单倍 | — |
| 关键词 | 楷体 | 小四 (12pt) | 否 | 左 | 单倍 | — |
| 一级标题 (一、) | 宋体 | 小四 (12pt) | **是** | 居中 | 单倍 | — |
| 二级标题 ((一)) | 楷体 | 小四 (12pt) | **是** | 左 | 单倍 | 缩进2字符 |
| 三级标题 (1. ) | 宋体 | 五号 (10.5pt) | **是** | 左 | 单倍 | 缩进2字符 |
| 正文 | 宋体 (含数字) | 五号 (10.5pt) | 否 | 左 | 单倍 | — |
| 脚注 | 宋体 | 小五号 (9pt) | 否 | — | 单倍 | 每页重编号, 1 2 3… |

## 触发

用户不需要记住斜杠命令 — 只要跟论文格式修改相关的自然语言请求都会激活此技能：

```
/修改论文格式的skill 格式化这篇论文 [文件路径]

使用修改论文格式skill，帮我把这篇论文排版调整成学术规范

用一下论文格式修改技能，处理这个文件夹里的所有论文

帮我把论文格式统一成中文规范

修改论文格式：标题黑体四号居中，正文宋体五号

这篇论文的排版需要规范化，帮我调整一下
```

**触发关键词组合**: 只要用户的请求中包含以下任一组合理念即可激活：
- 「论文」+「格式」/「排版」/「格式化」/「规范」
- 「修改」/「调整」/「统一」+「论文格式」/「论文排版」
- 明确提到「修改论文格式」或「论文格式修改」
- 直接引用 skill 名称：「使用修改论文格式 skill」「论文格式 skill」

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

### 4. 引注格式检查与修复

依据《法学引注手册》(2019) 规范，自动检查和修复脚注中的引注格式问题：

```bash
# 仅检查引注格式 (不改动)
python scripts/format_paper.py 论文.docx --check-citations

# 自动修复引注格式问题
python scripts/format_paper.py 论文.docx --fix-citations -o 输出.docx

# 完整处理: 排版格式化 + 引注修复
python scripts/format_paper.py 论文.docx --body-indent 2 --fix-citations -o 输出.docx
```

引注格式化涵盖:
- 中英文标点混用 → 统一为中文标点 (，、：；。)
- 作者名后逗号 → 改为冒号
- 英文页码格式 p.xx → 第xx页
- 出版社格式补全 (加"年版")
- "载"字缺失检测
- 书名号缺失检测

### 5. 独立引注检查工具

也可以单独使用引注格式化脚本：

```bash
python scripts/citation_formatter.py 论文.docx --check
python scripts/citation_formatter.py 论文.docx --fix -o 输出.docx
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

**自然触发模式** (不需要斜杠命令，说这些话就能激活):
- 「使用修改论文格式 skill」「用论文格式化技能」「调用论文格式修改」
- 「帮我调整论文排版」「把这篇论文格式改一下」「论文格式规范化」
- 「修改论文格式」「统一论文格式」「调整论文排版」
- 「format Chinese paper」「docx 论文格式

**实体**: 论文, 文章, 学位论文, 学术论文, 课程论文, docx, word文档
**动作**: 格式化, 排版, 调整格式, 修改格式, 规范化, 统一格式
**领域**: 中文学术规范, 论文排版, 学术写作, 毕业论文
