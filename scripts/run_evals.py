#!/usr/bin/env python3
"""Eval specification for chinese-paper-format-skill.

Defines the skill's loss function: binary checks + golden cases.
Format: autoresearch-universal compatible (rule 18).
"""

import json
import sys
from pathlib import Path

EVAL_SPEC = {
    "name": "chinese-paper-format-skill",
    "version": "1.0.0",
    "criteria": [
        {
            "id": "title-font",
            "description": "论文题目使用黑体(SimHei)、四号(14pt)、不加粗、居中",
            "type": "llm-judge",
            "prompt": "检查文档第一个非空段落: 字体是否为黑体/SimHei, 字号是否14pt, 是否不加粗, 是否居中对齐。报告 pass/fail 并说明原因。"
        },
        {
            "id": "heading-l1-format",
            "description": "一级标题 (一、) 格式正确: 宋体、小四(12pt)、加粗、居中、单倍行距",
            "type": "llm-judge",
            "prompt": "检查文档中以中文数字+顿号开头的段落(如 一、二、): 字体是否为宋体/SimSun, 字号12pt, 是否加粗, 是否居中对齐, 行距是否单倍。报告 pass/fail。"
        },
        {
            "id": "heading-l2-format",
            "description": "二级标题 ((一)) 格式正确: 楷体、小四(12pt)、加粗、缩进2字符",
            "type": "llm-judge",
            "prompt": "检查文档中以括号中文数字开头的段落(如 (一)(二)): 字体是否为楷体/KaiTi, 字号12pt, 是否加粗, 是否首行缩进约2字符。报告 pass/fail。"
        },
        {
            "id": "heading-l3-format",
            "description": "三级标题 (1. ) 格式正确: 宋体、五号(10.5pt)、加粗、缩进2字符",
            "type": "llm-judge",
            "prompt": "检查文档中以数字+点号+空格开头的短段落(如 1. 2.): 字体是否为宋体/SimSun, 字号10.5pt, 是否加粗, 是否首行缩进。报告 pass/fail。"
        },
        {
            "id": "body-font",
            "description": "正文内容中文字体宋体、西文字体Times New Roman、五号(10.5pt)、单倍行距",
            "type": "llm-judge",
            "prompt": "检查文档正文段落(非标题): 中文字体是否宋体/SimSun, 西文字体是否Times New Roman, 字号是否10.5pt, 行距是否单倍。报告 pass/fail。"
        },
        {
            "id": "footnote-format",
            "description": "脚注宋体小五号(9pt)、单倍行距、每页重新编号",
            "type": "llm-judge",
            "prompt": "如果文档有脚注: 检查脚注字体是否宋体/SimSun, 字号是否9pt/小五号。如果文档无脚注则 skip。报告 pass/fail/skip。"
        }
    ],
    "golden": [
        {
            "id": "sample-paper-minimal",
            "description": "最小化论文样本: 题目 + 一级标题 + 正文",
            "input": "evals/golden/sample-paper-minimal/input.docx",
            "expected_output": "evals/golden/sample-paper-minimal/expected.docx",
            "expected_status": "pending-first-green",
            "run": "python3 scripts/run_pipeline.py --input {input} --output {output} --no-check"
        },
        {
            "id": "sample-paper-full",
            "description": "完整论文样本: 题目 + 三级标题 + 正文 + 脚注",
            "input": "evals/golden/sample-paper-full/input.docx",
            "expected_output": "evals/golden/sample-paper-full/expected.docx",
            "expected_status": "pending-first-green",
            "run": "python3 scripts/run_pipeline.py --input {input} --output {output} --no-check"
        },
        {
            "id": "sample-paper-no-headings",
            "description": "纯正文无标题: 边界情况测试",
            "input": "evals/golden/sample-paper-no-headings/input.docx",
            "expected_output": "evals/golden/sample-paper-no-headings/expected.docx",
            "expected_status": "pending-first-green",
            "run": "python3 scripts/run_pipeline.py --input {input} --output {output} --no-check"
        }
    ]
}


def validate() -> bool:
    """Validate the eval spec itself is well-formed."""
    errors = []

    # Check required top-level keys
    for key in ('name', 'version', 'criteria', 'golden'):
        if key not in EVAL_SPEC:
            errors.append(f"Missing top-level key: {key}")

    # Check criteria
    for i, c in enumerate(EVAL_SPEC.get('criteria', [])):
        for field in ('id', 'description', 'type'):
            if field not in c:
                errors.append(f"criteria[{i}]: missing '{field}'")
        if c.get('type') not in ('command', 'llm-judge'):
            errors.append(f"criteria[{i}]: type must be 'command' or 'llm-judge'")
        if c['type'] == 'command' and 'cmd' not in c:
            errors.append(f"criteria[{i}]: command type requires 'cmd' field")
        if c['type'] == 'llm-judge' and 'prompt' not in c:
            errors.append(f"criteria[{i}]: llm-judge type requires 'prompt' field")

    # Check golden cases
    for i, g in enumerate(EVAL_SPEC.get('golden', [])):
        for field in ('id', 'description', 'input', 'expected_output', 'expected_status'):
            if field not in g:
                errors.append(f"golden[{i}]: missing '{field}'")

    if errors:
        print("VALIDATION FAILED:")
        for e in errors:
            print(f"  - {e}")
        return False

    print("VALID")
    print(f"  Criteria: {len(EVAL_SPEC['criteria'])} checks")
    print(f"  Golden cases: {len(EVAL_SPEC['golden'])}")
    return True


def main():
    if '--validate' in sys.argv:
        ok = validate()
        sys.exit(0 if ok else 1)
    elif '--spec' in sys.argv:
        print(json.dumps(EVAL_SPEC, ensure_ascii=False, indent=2))
    else:
        print("Usage: python run_evals.py --validate | --spec")
        sys.exit(1)


if __name__ == '__main__':
    main()
