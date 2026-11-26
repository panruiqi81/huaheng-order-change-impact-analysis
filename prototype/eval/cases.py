"""评估集：由基准快照派生的用例。tests 与 tools/run_eval.py 共用，保证测的和记录的是同一批。"""

from __future__ import annotations

import csv
import shutil
from dataclasses import dataclass, field
from pathlib import Path

BASE = Path(__file__).resolve().parents[1] / "data" / "base"
ORDER = "O-202610-001"

TABLES = {
    "erp": "erp_change_orders.csv",
    "mes": "mes_schedule.csv",
    "cap": "capacity.csv",
    "wms": "wms_inventory.csv",
    "qms": "qms_holds.csv",
}


def load_base():
    out = {}
    for key, name in TABLES.items():
        with (BASE / name).open("r", encoding="utf-8-sig", newline="") as fh:
            out[key] = list(csv.DictReader(fh))
    return out


def edit(t, key, index=0, **values):
    t[key][index].update(values)


@dataclass
class Case:
    id: str
    title: str
    expect: str                      # 预期总体结论
    codes: tuple = ()                # 必须出现的发现编号
    absent: tuple = ()               # 必须不出现的发现编号
    where: dict = field(default_factory=dict)   # 结果字段断言（点路径）
    setup: object = None             # 对基准表做修改的函数
    path: str = "正常路径"
    facts: str = ""
    order: str = ORDER


def _released_v3(t):
    edit(t, "erp", 1, status="released", countersign_plan="yes", countersign_sales="yes",
         released_by="王建国/陈思雨", released_at="2026-11-04T10:00:00")


def _b17(t, wms_status):
    t["wms"].append({"order_id": ORDER, "sku": "SKU-A1", "batch_id": "B-17", "available_qty": "200",
                     "wms_hold_status": wms_status, "recorded_at": "2026-11-03T09:00:00"})


def _drop_table(t, key):
    t[key].clear()


def _drop_file(t, key):
    t.setdefault("_removed", []).append(key)


def _drop_mes(t):
    _drop_table(t, "mes")


CASES = [
    Case("TC-01", "基准：V3 已提出未会签，MES 按 V2 排产，缺口 100 件", "ATTENTION",
         ("PENDING_CHANGE", "SCHEDULE_ALIGNED", "OVERTIME_PENDING", "CAPACITY_GAP"), ("NO_APPROVED_VERSION", "BATCH_HELD"),
         {"capacity_check.confirmed_capacity": 900, "capacity_check.usable_inventory": 200, "capacity_check.pending_overtime": 150},
         path="正常路径", facts="f12 / f20 / f21 / f22 / f23"),
    Case("TC-02", "V3 已批准，MES 仍按 V2 排产（现场在做旧版）", "CONFLICT", ("SCHEDULE_BEHIND", "CAPACITY_GAP"), ("SCHEDULE_ALIGNED",),
         {"version_check.approved_version": "V3", "capacity_check.scenarios.1.gap": 100}, _released_v3,
         path="异常路径 A · 版本冲突", facts="f20 / f25"),
    Case("TC-03", "V3 未批准，MES 已按 V3 排产（超前排产）", "CONFLICT", ("SCHEDULE_AHEAD", "PENDING_CHANGE"), ("SCHEDULE_ALIGNED",),
         {"capacity_check.scenarios.1.gap": 100}, lambda t: edit(t, "mes", planned_version="V3", qty="1200"),
         path="异常路径 B · 未批先排", facts="f11 / f25"),
    Case("TC-04", "订单没有任何已批准版本", "CONFLICT", ("NO_APPROVED_VERSION", "SCHEDULE_WITHOUT_APPROVAL"), ("SCHEDULE_ALIGNED",),
         {"version_check.approved_version": None},
         lambda t: edit(t, "erp", 0, status="proposed", countersign_plan="no", countersign_sales="no", released_by="", released_at=""),
         path="异常路径 A · 版本冲突", facts="f14"),
    Case("TC-05", "V3 会签完成但尚未签发，不算批准", "ATTENTION", ("PENDING_CHANGE", "SCHEDULE_ALIGNED"), ("SCHEDULE_BEHIND",),
         {"version_check.approved_version": "V2"},
         lambda t: edit(t, "erp", 1, status="countersigned", countersign_plan="yes", countersign_sales="yes"),
         path="正常路径", facts="f14 / f23"),
    Case("TC-06", "V3 已批准且 MES 已切换，产能充足", "OK", ("SCHEDULE_ALIGNED",), ("CAPACITY_GAP", "PENDING_CHANGE", "OVERTIME_PENDING"),
         {"capacity_check.scenarios.1.gap": 0, "capacity_check.scenarios.0.gap": 0, "pending": []},
         lambda t: (_released_v3(t), edit(t, "mes", planned_version="V3", qty="1200"), edit(t, "cap", confirmed_capacity="1000", pending_overtime="0")),
         path="正常路径", facts="f21"),
    Case("TC-07", "待批加班 300 件足以覆盖缺口，但未获批不得计入产能", "ATTENTION", ("CAPACITY_GAP", "OVERTIME_PENDING"), (),
         {"capacity_check.scenarios.1.gap": 100, "capacity_check.pending_overtime": 300, "capacity_check.confirmed_capacity": 900},
         lambda t: edit(t, "cap", pending_overtime="300"),
         path="异常路径 C · 缺口依赖加班", facts="f22"),
    Case("TC-08", "订单占用冻结批次 B-17（QMS 与 WMS 一致）", "HARD_STOP", ("BATCH_HELD",), ("QMS_WMS_MISMATCH",),
         {"capacity_check.usable_inventory": 200, "capacity_check.scenarios.1.gap": 100}, lambda t: _b17(t, "held"),
         path="异常路径 D · 质量冻结", facts="f40 / f43"),
    Case("TC-09", "B-17 冻结，但 WMS 仍显示已放行（同步延迟）", "HARD_STOP", ("BATCH_HELD", "QMS_WMS_MISMATCH"), (),
         {"capacity_check.usable_inventory": 200}, lambda t: _b17(t, "released"),
         path="异常路径 E · 冻结状态同步延迟", facts="f41"),
    Case("TC-10", "批次返工完成但未重新检验", "HARD_STOP", ("REWORK_NOT_INSPECTED",), ("BATCH_HELD",),
         {"capacity_check.usable_inventory": 0, "capacity_check.scenarios.1.gap": 300},
         lambda t: edit(t, "qms", 0, hold_status="rework_pending_inspection", record_version="3", signed_by=""),
         path="异常路径 D · 质量冻结", facts="f42"),
    Case("TC-11", "放行记录没有签发人，不视为有效放行", "HARD_STOP", ("RELEASE_UNSIGNED",), ("BATCH_HELD",),
         {"capacity_check.usable_inventory": 0}, lambda t: edit(t, "qms", 0, signed_by=""),
         path="异常路径 D · 质量冻结", facts="f40 / f45"),
    Case("TC-12", "批次在 QMS 中没有质量记录", "MANUAL_REVIEW", ("NO_QUALITY_RECORD",), ("BATCH_HELD",),
         {"capacity_check.usable_inventory": 0}, lambda t: edit(t, "wms", 0, batch_id="B-18"),
         path="异常路径 F · 数据缺失转人工", facts="f32 / f44"),
    Case("TC-13", "订单在 WMS 没有库存记录（缺记录 ≠ 零库存）", "MANUAL_REVIEW", ("NO_INVENTORY",), (),
         {"capacity_check.usable_inventory": 0, "capacity_check.scenarios.1.gap": 300}, lambda t: _drop_table(t, "wms"),
         path="异常路径 F · 数据缺失转人工", facts="f32"),
    Case("TC-14", "缺少产能快照文件", "MANUAL_REVIEW", (), ("CAPACITY_GAP",), {"version_check": None, "capacity_check": None},
         lambda t: _drop_file(t, "cap"), path="异常路径 F · 数据缺失转人工", facts="f32"),
    Case("TC-15", "变更单数量字段不是整数（脏数据）", "MANUAL_REVIEW", (), ("CAPACITY_GAP",), {"version_check": None},
         lambda t: edit(t, "erp", 1, qty="约1200"), path="异常路径 F · 数据缺失转人工", facts="f32"),
    Case("TC-16", "MES 没有该订单的排产记录", "MANUAL_REVIEW", ("NO_SCHEDULE", "NO_CAPACITY"), ("SCHEDULE_ALIGNED",), {},
         _drop_mes, path="异常路径 F · 数据缺失转人工", facts="f20 / f32"),
    Case("TC-17", "查询不存在的订单号", "MANUAL_REVIEW", ("NO_CHANGE_ORDER",), (), {},
         None, path="异常路径 F · 数据缺失转人工", facts="f31", order="O-999999"),
]

BY_ID = {c.id: c for c in CASES}


def materialize(case, dest):
    """把用例写成一组快照文件。返回快照目录。"""
    dest = Path(dest)
    if dest.exists():
        shutil.rmtree(dest)
    shutil.copytree(BASE, dest)
    t = load_base()
    if case.setup:
        case.setup(t)
    for key, name in TABLES.items():
        path = dest / name
        if key in t.get("_removed", []):
            path.unlink()
            continue
        with path.open("w", encoding="utf-8", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=list(load_base()[key][0].keys()))
            w.writeheader()
            w.writerows(t[key])
    return dest
