"""跑完整评估集并把真实结果写成 eval/results.md。

用法（在 prototype/ 目录下）：python tools/run_eval.py
"""

from __future__ import annotations

import hashlib
import platform
import sys
import tempfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from eval.cases import BY_ID, CASES, ORDER, TABLES, materialize  # noqa: E402
from eval.verify import check, codes_of, get  # noqa: E402
from oc_checker import RULES_VERSION  # noqa: E402
from oc_checker.__main__ import run  # noqa: E402
from oc_checker.render import to_markdown  # noqa: E402

# 评估使用固定时间戳，保证 results.md 可复现
FIXED_NOW = datetime(2026, 11, 26, 4, 0, tzinfo=timezone.utc)


def digest_dir(path):
    return {name: hashlib.sha256((Path(path) / name).read_bytes()).hexdigest()
            for name in TABLES.values() if (Path(path) / name).is_file()}


def main():
    rows, base_report = [], None
    with tempfile.TemporaryDirectory() as tmp:
        for case in CASES:
            snap = materialize(case, Path(tmp) / case.id)
            before = digest_dir(snap)
            result = run(snap, case.order, now=FIXED_NOW)
            after = digest_dir(snap)
            fails = check(case, result)
            if before != after:
                fails.append("运行后输入快照发生变化（违反只读约束）")
            if case.id == "TC-01":
                base_report = to_markdown(result)
            rows.append((case, result, fails))

    passed = sum(1 for _, _, f in rows if not f)
    by_path = Counter(c.path for c, _, _ in rows)
    by_status = Counter(r["overall_status"] for _, r, _ in rows)

    out = [
        "# 评估结果 · 订单变更影响核对助手（只读原型）", "",
        f"- 运行时间戳：{FIXED_NOW.isoformat()}（评估固定值，便于复现）",
        f"- 规则版本：`{RULES_VERSION}` · Python {platform.python_version()} · {platform.system()} {platform.release()}",
        f"- 用例数：{len(rows)} · 通过：**{passed}** · 失败：{len(rows) - passed}",
        f"- 结论分布：{'、'.join(f'{k} {v}' for k, v in sorted(by_status.items()))}",
        f"- 路径分布：{'、'.join(f'{k} {v}' for k, v in sorted(by_path.items()))}",
        "",
        "| 用例 | 场景 | 路径 | 预期 | 实际 | 实际命中的发现编号 | 结果 |",
        "|---|---|---|---|---|---|---|",
    ]
    for case, result, fails in rows:
        out.append(f"| {case.id} | {case.title} | {case.path.split(' · ')[-1]} | {case.expect} | {result['overall_status']} "
                   f"| {', '.join(codes_of(result)) or '—'} | {'✅ 通过' if not fails else '❌ ' + '；'.join(fails)} |")
    out += ["", "## 评估口径", "",
            "1. 每个用例由基准快照（`data/base/`）派生，改动项写在 `eval/cases.py`，可逐条复现。",
            "2. 通过条件：总体结论、必须出现/不得出现的发现编号、关键字段值全部符合预期，且运行前后输入快照哈希不变（只读）。",
            "3. 预期值是先写后测的：先按访谈事实定下每个用例“应该得出什么”，再运行比对；失败项原样保留，不改预期去迁就结果。",
            "4. 评估对象是规则引擎的确定性输出，不涉及大模型；接入模型生成说明文字后需另建评估集。", ""]
    (ROOT / "eval" / "results.md").write_text("\n".join(out), encoding="utf-8")
    (ROOT / "eval" / "base-report.md").write_text(base_report, encoding="utf-8")
    print("\n".join(out[:9]))
    for case, _, fails in rows:
        if fails:
            print(case.id, "FAIL:", "；".join(fails))
    return 0 if passed == len(rows) else 1


if __name__ == "__main__":
    sys.exit(main())
