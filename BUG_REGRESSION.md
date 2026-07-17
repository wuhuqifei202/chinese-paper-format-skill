# Bug 回归清单

每个已修复的 bug 对应一个编号。新增 bug 修复时必须同步更新此文件。

格式：
```
## BUG-XXX: 标题
- **修复**: <commit-hash>
- **原因**: <一句话描述根因>
- **测试**: <空格分隔的测试函数名>
```

---

## BUG-001: 脚注文本出现在自动编号之前

- **修复**: de40e85
- **原因**: format_all_footnotes 将修复后文本写入第一个 w:t 元素，覆盖了 footnoteRef 之前的括号，导致自动编号跑到文本后面
- **测试**: test_ref_before_text test_ref_preserved_after_fix test_citation_text_after_fix

## BUG-002: 摘要/关键词被擅自添加 2 字符缩进

- **修复**: d0d1858
- **原因**: 实现 is_abstract_or_keywords 时未逐条确认用户需求，自作主张加了 set_first_line_indent(2)
- **测试**: test_abstract_no_indent test_abstract_kaiti test_keywords_kaiti

## BUG-003: ZIP 降级函数序列化了修改前的 XML

- **修复**: d2c110b (第二次提交)
- **原因**: write_footnotes_via_zipfile 在调用 format_all_footnotes 前获取了 fn_xml 引用，修复后未重新解析，序列化了旧 XML
- **测试**: test_blob_write_back test_citation_fixes_applied

## BUG-004: _blob 写回分散在两处，无错误处理

- **修复**: d2c110b
- **原因**: format_paper.py 和 citation_formatter.py 各自直接访问 _blob，无版本检测、无降级方案
- **测试**: test_blob_write_back test_citation_fixes_applied

## BUG-005: 阿拉伯数字使用 Times New Roman 而非宋体

- **修复**: ac0c2ba
- **原因**: FONT_TNR 常量设为 Times New Roman，ascii/hAnsi 字体指向 TNR，导致数字也渲染为 TNR
- **测试**: test_ascii_font_is_song

## BUG-006: 摘要和关键词被格式化为黑体四号（与标题相同）

- **修复**: 44afac6
- **原因**: 标题区域检测将第一标题前的所有段落统一格式化为标题样式，未区分摘要/关键词
- **测试**: test_abstract_kaiti test_keywords_kaiti test_abstract_no_indent
