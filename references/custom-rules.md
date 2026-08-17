# 自定义格式规则

默认情况下 skill 按预设格式排版（题目黑体四号居中、正文宋体五号…）。
通过 `--rules 配置.json` 可以按自己的格式要求便捷转换格式。

## 用法

```bash
# 直接格式化 (工作流 1)
python scripts/format_paper.py 论文.docx --rules 我的格式.json -o 输出.docx

# markdown 中转工作流 (双交付)
python scripts/run_pipeline.py --input 论文.docx --via-markdown --rules 我的格式.json

# 批量
python scripts/format_paper.py 论文/ --batch --rules 我的格式.json
```

## 配置文件格式

JSON 文件，**只写要改的元素**（未写出的元素/字段用默认规则），
`_` 开头的键视为注释可以随便写：

```json
{
  "_说明": "示例: 题目二号加粗, 正文小四 20磅行距",
  "title": {"size": "二号", "bold": true},
  "body":  {"size": "小四", "line_spacing": "20磅"}
}
```

完整示例见 `assets/rules.example.json`（复制修改即可用）。

## 支持的元素

| 键 | 含义 |
|---|---|
| `title` | 论文题目 |
| `author` | 作者行 / 副标题（题目区域非首段非元数据） |
| `abstract` | 摘要 |
| `keywords` | 关键词 |
| `h1` | 一级标题（一、） |
| `h2` | 二级标题（（一）） |
| `h3` | 三级标题（1. ） |
| `h4` | 四级标题（（1）） |
| `body` | 正文 |
| `footnote` | 脚注 |

## 支持字段

| 字段 | 说明 | 示例 |
|---|---|---|
| `font` | 中文字体名 | `"黑体"` `"宋体"` `"楷体"` `"仿宋"` |
| `size` | 字号：中文名 / 磅值数字 / "22pt" | `"二号"` `22` `"22pt"` |
| `bold` | 加粗 | `true` `false` |
| `align` | 对齐 | `"center"` `"left"` `"right"` `"justify"`（或 `"居中"` `"左对齐"`…） |
| `indent` | 首行缩进字符数（0=无） | `2` `0` |
| `line_spacing` | 行距：数字=倍数，"N磅"=固定值 | `1.5` `"20磅"` `"单倍"` |

## 中文字号表

`初号42 小初36 一号26 小一24 二号22 小二18 三号16 小三15 四号14 小四12 五号10.5 小五9 六号7.5 小六6.5 七号5.5 八号5`（单位: 磅）

## 示例

某期刊要求：题目黑体二号加粗；一级标题宋体四号加粗；二级标题小四宋体加粗；
三级四级标题小四宋体；正文小四宋体行距 20 磅：

```json
{
  "title":  {"size": "二号", "bold": true},
  "h1":     {"font": "宋体", "size": "四号", "bold": true},
  "h2":     {"font": "宋体", "size": "小四", "bold": true},
  "h3":     {"font": "宋体", "size": "小四", "bold": false},
  "h4":     {"font": "宋体", "size": "小四", "bold": false},
  "body":   {"size": "小四", "line_spacing": "20磅"}
}
```

## 校验规则（加载时报错并给出提示）

- 未知元素 / 未知字段 → 报错
- 未知字号（如"特大"）→ 报错
- 未知对齐方式 → 报错
- 配置文件 JSON 解析失败 → 报错

配置出错时脚本输出友好错误并退出（不修改文档），可先用
`python scripts/rules.py 配置.json` 单独校验。
