"""用例判定：收集发现编号、按点路径取字段、比对预期。run_eval 与 tests 共用。"""

from __future__ import annotations


def codes_of(result):
    codes = []
    for key in ("version_check", "quality_findings", "capacity_check"):
        block = result.get(key)
        if not block:
            continue
        if isinstance(block, list):          # quality_findings 本身就是发现列表
            findings = block
        else:
            findings = block.get("findings", [])
        for f in findings:
            codes.append(f["code"])
    return codes


def get(result, path):
    cur = result
    for part in path.split("."):
        cur = cur[int(part)] if isinstance(cur, list) else cur[part]
    return cur


def check(case, result):
    """返回失败说明列表；为空表示通过。"""
    fails = []
    if result["overall_status"] != case.expect:
        fails.append(f"结论应为 {case.expect}，实际 {result['overall_status']}")
    codes = codes_of(result)
    for c in case.codes:
        if c not in codes:
            fails.append(f"缺少发现 {c}")
    for c in case.absent:
        if c in codes:
            fails.append(f"不应出现发现 {c}")
    for path, want in case.where.items():
        got = get(result, path)
        if got != want:
            fails.append(f"{path} 应为 {want!r}，实际 {got!r}")
    return fails
