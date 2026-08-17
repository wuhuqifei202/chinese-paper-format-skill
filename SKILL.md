---
name: chinese-paper-format-skill
description: >-
  将中文论文 doc/docx 按学术排版规范自动格式化: 题目(黑体四号居中)、
  一/二/三/四级标题(宋体/楷体小四·五号加粗缩进)、正文(宋体五号单倍行距缩进2)、
  脚注(宋体小五号每页重编号)、引注规范化(《法学引注手册》2019)。
  支持 markdown 中转工作流 (--via-markdown): doc/docx 先转结构化 md 再重建
  .docx, md 与 docx 双交付; .doc 旧格式用 LibreOffice 自动转换。
  支持自定义格式规则 (--rules 配置.json), 按自己的格式要求便捷转换
  (如题目黑体二号加粗、正文小四20磅行距), 未指定元素用默认规则。
  支持自定义注释体例: 用户在对话框用文字描述注释/引注要求即可, 无需手写
  JSON (内部转为 --citation-rules 配置.json 或 --preset 预设); 用户未提出
  要求时按默认体例(法学引注手册)修改。
  当用户说"修改论文格式"、"论文格式修改"、"帮我调整论文排版"、
  "格式化这篇论文"、"统一论文格式"、"docx论文格式调整"
  或任何与论文格式修改相关的请求时激活此技能。
license: MIT
metadata:
  author: Claude Code Agent Skill Creator
  version: 1.3.0
  created: 2026-07-16
  last_reviewed: 2026-08-17
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
  version: 1.3.0
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
| 四级标题 ((1)) | 宋体 | 五号 (10.5pt) | **是** | 左 | 单倍 | 缩进2字符 |
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

默认依据《法学引注手册》(2019) 自动检查和修复脚注中的引注格式问题（用户未
提出要求时即用此默认体例）：

```bash
# 仅检查引注格式 (不改动)
python scripts/format_paper.py 论文.docx --check-citations

# 自动修复引注格式问题
python scripts/format_paper.py 论文.docx --fix-citations -o 输出.docx

# 完整处理: 排版格式化 + 引注修复
python scripts/format_paper.py 论文.docx --body-indent 2 --fix-citations -o 输出.docx

# 按自定义注释体例修复 (配置由 skill 根据用户文字自动生成)
python scripts/format_paper.py 论文.docx --fix-citations --citation-rules 体例.json -o 输出.docx
```

引注格式化涵盖:
- 中英文标点混用 → 统一为中文标点 (，、：；。)
- 作者名后逗号 → 改为冒号
- 英文页码格式 p.xx → 第xx页
- 出版社格式补全 (加"年版")
- "载"字缺失检测
- 书名号缺失检测
- 文号方括号 → 六角括号、案号方括号 → 圆括号

可选（默认关闭，用户要求时启用）:
- 版次归位: 《书名》第N版 → 《书名》（第N版）
- 译者名与"译"之间的空格清理

### 5. 独立引注检查工具

也可以单独使用引注格式化脚本：

```bash
python scripts/citation_formatter.py 论文.docx --check
python scripts/citation_formatter.py 论文.docx --fix -o 输出.docx
python scripts/citation_formatter.py --show-style                          # 打印默认体例(法学引注手册)
python scripts/citation_formatter.py --show-style --preset 《法学家》《中外法学》注释体例  # 打印两刊预设
python scripts/citation_formatter.py --citation-rules 体例.json --show-style
```

### 6. Markdown 中转工作流 (推荐)

输入阶段先将 doc/docx 转换为结构化 markdown，再进行修改与重建。md 只承载
**内容与结构** (段落/标题层级/脚注)，排版在重建 .docx 时按上方规则表统一套用；
输出为 **md 与 docx 双交付** (修改后的 .md + 重建的 .docx)。

```bash
# 一键中转: docx → md → (引注修复) → 重建 docx, 输出 .md + .docx
python scripts/run_pipeline.py --input 论文.docx --via-markdown --fix-citations

# 旧版 .doc 自动转换 (需要 LibreOffice, 未安装时给出提示)
python scripts/run_pipeline.py --input 论文.doc --via-markdown

# 分步使用 (可对 md 做任何审查/修改后再重建)
python scripts/docx2md.py 论文.docx -o 论文.md        # ① docx → md
# ... 人工/自动修改 论文.md (引注修复、结构调整) ...
python scripts/md2docx.py 论文.md -o 论文_已格式化.docx  # ② md → 重建 docx (排版自动套用)
```

**markdown 结构约定** (docx2md 输出 / md2docx 输入，二者互逆):

```
# 论文题目                 → 题目 (黑体四号居中)
作者行 / 副标题            → 题目区域其他段 (黑体四号居中)
【摘要】... 【关键词】...  → 楷体小四
## 一、一级标题
### （一）二级标题
#### 1. 三级标题
##### （1）四级标题
正文段落[^1]               → 正文 (宋体五号缩进2)
[^1]: 脚注文本             → 脚注定义 (文末, 按编号升序)
```

**适用场景**: 需要审查/修改论文内容 (引注、结构、文字) 后再统一排版；
不需要人工介入时可直接用工作流 1 (原地排版)。

**限制**: md 不承载排版格式 (字体/字号/对齐) 与复杂对象 (表格/图片/批注) —
这些在 docx→md 中丢失，重建时按规则表重新套用；含表格/图片的论文建议使用
工作流 1。

### 7. 自定义格式规则 (按需换格式)

默认排版不够时，用 `--rules 配置.json` 按自己的格式要求便捷转换，
**只写要改的元素**，未写出的元素/字段用默认规则。示例: 某期刊要求
题目黑体二号加粗、一级标题宋体四号加粗、二级标题小四宋体加粗、
三级四级标题小四宋体、正文小四宋体行距20磅:

```json
{
  "title": {"size": "二号", "bold": true},
  "h1":    {"font": "宋体", "size": "四号", "bold": true},
  "h2":    {"font": "宋体", "size": "小四", "bold": true},
  "h3":    {"font": "宋体", "size": "小四", "bold": false},
  "h4":    {"font": "宋体", "size": "小四", "bold": false},
  "body":  {"size": "小四", "line_spacing": "20磅"}
}
```

```bash
# 直接格式化 (工作流 1) + 自定义规则
python scripts/format_paper.py 论文.docx --rules 我的格式.json -o 输出.docx

# markdown 中转工作流 + 自定义规则 (md 与 docx 双交付)
python scripts/run_pipeline.py --input 论文.docx --via-markdown --rules 我的格式.json

# 批量
python scripts/format_paper.py 论文/ --batch --rules 我的格式.json
```

支持元素: `title` `author` `abstract` `keywords` `h1`–`h4` `body` `footnote`;
字段: `font`(中文字体) `size`(中文名/磅值) `bold` `align`(center/left/right/justify)
`indent`(首行缩进字符数) `line_spacing`(数字=倍数, "20磅"=固定磅值)。
`_` 开头的键视为注释可写在配置里; 配置出错时脚本给出友好错误并不修改文档。
完整说明见 `references/custom-rules.md`, 示例文件 `assets/rules.example.json`。

### 8. 自定义注释格式（注释体例）

默认依据《法学引注手册》(2019) 检查/修复脚注引注。**用户未提出自己的要求时，
按默认体例修改**；用户提出自己的注释体例时，优先按用户要求转换。

**用户不需要自己写 JSON** — 直接在对话框用文字描述即可，由 skill 翻译成配置：

- 用户说「用《法学家》《中外法学》的注释体例（文献引用格式）」→ `--preset 《法学家》《中外法学》注释体例`
- 用户说「用《中国法学》/《法商研究》的引注体例」→ `--preset 《中国法学》` / `--preset 《法商研究》`（两者与法学引注手册一致）
- 用户说「用《法学研究》的注释体例」→ `--preset 《法学研究》`（法学引注手册 + 14 条要求）
- 用户说「版次放括号里：《书名》（第N版）」→ 生成 `{"rules": {"edition": true}}`
- 用户贴出完整体例（如「著作类：作者《书名》，出版社年版，第N页；论文类：……」）→
  逐条翻译为 `categories` 模板 + 相应 `rules` 开关，写入临时 JSON 后传入

```bash
# 查看默认体例 (法学引注手册) / 各期刊预设
python scripts/citation_formatter.py --show-style
python scripts/citation_formatter.py --show-style --preset 《法学家》《中外法学》注释体例
python scripts/citation_formatter.py --show-style --preset 《法学研究》

# 直接格式化 + 引注修复, 指定预设或体例
python scripts/format_paper.py 论文.docx --fix-citations --preset 《法学研究》 -o 输出.docx
python scripts/format_paper.py 论文.docx --fix-citations --citation-rules 体例.json -o 输出.docx

# markdown 中转工作流 + 自定义体例
python scripts/run_pipeline.py --input 论文.docx --via-markdown --fix-citations --citation-rules 体例.json

# 校验并预览体例
python scripts/citation_rules.py 体例.json
```

配置只写要改的部分（`_` 开头的键为注释），可配置维度：
`rules`(各修复规则开关)、`conventions`(命名约定)、`categories`(各类体例模板/示例)、
`notes`(体例要求/注意事项，如《法学研究》的 14 条)。
默认体例与预设、字段说明见 `references/citation-rules.md`，示例文件
`assets/citation_rules.example.json`。

## 自动检测机制

脚本通过内容模式自动识别文档结构，**不需要手动标记**:

- **论文题目**: 文档第一个非空段落
- **一级标题**: 匹配 `一、` `二、` … `十一、` … 等中文数字编号，且段落较短
- **二级标题**: 匹配 `（一）` `（二）` … 等括号中文数字编号
- **三级标题**: 匹配 `1. ` `2. ` … 等数字编号后带空格
- **四级标题**: 匹配 `（1）` `（2）` … 等括号阿拉伯数字编号，且段落较短 (≤20字符)
- **正文**: 不匹配任何标题模式的段落
- **脚注**: 文档中已有的 Word 脚注

## 注意事项

1. **.docx 直接支持**; **.doc 旧格式** 仅在 `--via-markdown` 模式下自动转换 (需安装
   [LibreOffice](https://www.libreoffice.org/), 未安装时给出安装提示; 也可先用 Word
   另存为 .docx)
2. **覆盖原文件自动备份** (格式: `原文件名.backup_时间戳.docx`)
3. **脚注只格式化已有的**。如果原文中注释是手动输入的文本 (非 Word 脚注功能), 需要用户先转换为 Word 脚注
4. **正文首行缩进默认 2 字符**。可用 `--body-indent 0` 关闭
5. **markdown 中转丢失排版与复杂对象** (表格/图片/批注)。需要保留这些内容时用工作流 1 (原地排版)

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
