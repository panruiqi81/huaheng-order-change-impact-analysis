"""命令行入口：python -m oc_checker --snapshot DIR --order ORDER [--out DIR]"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from pathlib import Path

from .audit import bind_task, write_manifest
from .constants import LEVEL_LABEL
from .models import load_snapshot
from .render import to_json, to_markdown
from .rules import evaluate


def run(snapshot_dir, order_id, out_dir=None, now=None):
    now = now or datetime.now(timezone.utc)
    snap = load_snapshot(snapshot_dir)
    result = evaluate(snap, order_id)
    result["task"] = bind_task(snap, order_id, now)
    if out_dir:
        out = Path(out_dir)
        out.mkdir(parents=True, exist_ok=True)
        md, js = out / "report.md", out / "report.json"
        md.write_text(to_markdown(result), encoding="utf-8")
        js.write_text(to_json(result), encoding="utf-8")
        write_manifest(result["task"], [md, js])
    return result


def main(argv=None):
    ap = argparse.ArgumentParser(prog="oc_checker", description="订单变更影响核对助手（只读）")
    ap.add_argument("--snapshot", required=True, help="快照目录（五张 CSV）")
    ap.add_argument("--order", required=True, help="订单号，如 O-202610-001")
    ap.add_argument("--out", help="输出目录：写入 report.md / report.json / manifest.json")
    a = ap.parse_args(argv)
    result = run(a.snapshot, a.order, a.out)
    print(f"{result['task']['task_id']}  {LEVEL_LABEL[result['overall_status']]}")
    if not a.out:
        sys.stdout.write(to_markdown(result))
    return 0


if __name__ == "__main__":
    sys.exit(main())
