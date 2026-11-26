"""审计绑定：规则版本、输入快照与输出共用同一任务编号。"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

from .constants import RULES_VERSION


def input_digest(snap, order_id):
    h = hashlib.sha256()
    h.update(RULES_VERSION.encode())
    h.update(order_id.encode())
    for name in sorted(snap.file_hashes):
        h.update(name.encode())
        h.update(snap.file_hashes[name].encode())
    return h.hexdigest()


def bind_task(snap, order_id, generated_at):
    """generated_at: datetime（含时区）。任务编号 = 日期 + 输入摘要前 8 位。"""
    digest = input_digest(snap, order_id)
    return {
        "task_id": f"HH-{generated_at:%Y%m%d}-{digest[:8]}",
        "input_digest": digest,
        "rules_version": RULES_VERSION,
        "order_id": order_id,
        "generated_at": generated_at.isoformat(timespec="seconds"),
        "input_files": dict(sorted(snap.file_hashes.items())),
    }


def write_manifest(task, paths):
    """把输出文件的哈希写进清单，形成 输入 → 任务编号 → 输出 的闭环。"""
    manifest = dict(task)
    manifest["outputs"] = {Path(p).name: hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in paths}
    manifest_path = Path(paths[0]).parent / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest_path
