#!/usr/bin/env python3
"""
Bug 回归清单校验工具

解析 BUG_REGRESSION.md，收集每个 bug 对应的测试函数名，
用 pytest --co 确认所有测试仍然存在。若有 bug 的测试被删除，报错退出。

Usage:
    python scripts/check_regression.py          # 检查所有 bug 的测试存在性
    python scripts/check_regression.py --list    # 列出所有 bug 和对应测试
    python scripts/check_regression.py --strict # 测试存在 + 必须通过
"""

import re
import sys
import subprocess
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
BUG_FILE = SKILL_DIR / 'BUG_REGRESSION.md'


def parse_regression_file() -> list[dict]:
    """解析 BUG_REGRESSION.md，返回 bug 列表."""
    if not BUG_FILE.exists():
        print(f"错误: {BUG_FILE} 不存在")
        sys.exit(1)

    content = BUG_FILE.read_text(encoding='utf-8')
    bugs = []

    # 匹配 ## BUG-XXX ... ## BUG-YYY 之间的内容
    bug_blocks = re.findall(
        r'## (BUG-\d+): (.+?)\n(.*?)(?=\n## BUG-|\Z)',
        content, re.DOTALL
    )

    for bug_id, title, body in bug_blocks:
        # 提取测试名
        test_match = re.search(r'\*\*测试\*\*:\s*(.+)', body)
        test_names = []
        if test_match:
            test_names = test_match.group(1).strip().split()

        # 提取 commit hash
        commit_match = re.search(r'\*\*修复\*\*:\s*(\S+)', body)
        commit = commit_match.group(1) if commit_match else 'unknown'

        bugs.append({
            'id': bug_id,
            'title': title.strip(),
            'commit': commit,
            'tests': test_names,
        })

    return bugs


def collect_all_tests() -> set[str]:
    """用 pytest --co 收集所有测试函数名."""
    result = subprocess.run(
        [sys.executable, '-m', 'pytest', 'tests/', '--co', '-q'],
        capture_output=True, text=True, cwd=SKILL_DIR
    )
    output = result.stdout + result.stderr
    # --co 输出 XML 风格: <Function test_name>
    tests = set()
    for match in re.finditer(r'<Function\s+([^>]+)>', output):
        tests.add(match.group(1))
    return tests


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Bug 回归清单校验')
    parser.add_argument('--list', action='store_true', help='列出所有 bug')
    parser.add_argument('--strict', action='store_true',
                        help='不仅检查测试存在，还要求全部通过')
    args = parser.parse_args()

    bugs = parse_regression_file()
    all_tests = collect_all_tests()

    if args.list:
        print(f"{'ID':<10} {'测试数':<6} 标题")
        print("-" * 60)
        for b in bugs:
            print(f"{b['id']:<10} {len(b['tests']):<6} {b['title']}")
            for t in b['tests']:
                exists = '✓' if t in all_tests else '✗ MISSING'
                print(f"           {exists}  {t}")
            print()
        return

    # 检查每个 bug 的测试是否存在
    missing = []
    for b in bugs:
        for t in b['tests']:
            if t not in all_tests:
                missing.append((b['id'], t))

    if missing:
        print("=" * 60)
        print("  ✗ 回归保护失效 — 以下测试已被删除:")
        print("=" * 60)
        for bug_id, test_name in missing:
            print(f"  {bug_id}: {test_name}")
        print()
        print(f"共 {len(missing)} 个测试缺失。")
        print("请恢复被删除的测试，或更新 BUG_REGRESSION.md 移除对应条目。")
        sys.exit(1)

    # strict 模式：运行指定测试
    if args.strict:
        all_needed = []
        for b in bugs:
            all_needed.extend(b['tests'])
        if not all_needed:
            print("✓ 无测试需要运行")
            return

        # 用 -k 运行所有回归测试
        expr = ' or '.join(all_needed)
        result = subprocess.run(
            [sys.executable, '-m', 'pytest', 'tests/', '-k', expr, '-q', '--tb=line'],
            cwd=SKILL_DIR
        )
        if result.returncode != 0:
            print("\n✗ 回归测试未全部通过")
            sys.exit(result.returncode)

    print(f"✓ 所有 {len(bugs)} 个 bug 的回归测试均存在 "
          f"({sum(len(b['tests']) for b in bugs)} 个测试函数)")
    for b in bugs:
        print(f"  {b['id']}: {len(b['tests'])} 测试 → {b['title']}")


if __name__ == '__main__':
    main()
