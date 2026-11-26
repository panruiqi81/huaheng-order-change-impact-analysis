"""快照加载与校验。缺文件、缺字段、非法取值都记入 Snapshot.errors，不抛出、不猜测。"""

from __future__ import annotations

import csv
import hashlib
from dataclasses import dataclass, field
from pathlib import Path

FILES = (
    "erp_change_orders.csv",
    "mes_schedule.csv",
    "capacity.csv",
    "wms_inventory.csv",
    "qms_holds.csv",
)

ERP_STATUSES = {"proposed", "countersigned", "released"}
HOLD_STATUSES = {"held", "rework_pending_inspection", "released"}
YES_NO = {"yes", "no"}
OPT = "optional"

# 这些字段允许为空：未签发的版本没有签发人和签发时间；放行记录缺签发人由规则层判为硬约束中止
OPTIONAL_FIELDS = {"released_by", "released_at", "signed_by"}


class InputError(Exception):
    """快照无法读取时的异常，仅用于编程错误；业务上的脏数据记入 Snapshot.errors。"""


@dataclass(frozen=True)
class ChangeOrder:
    order_id: str
    change_version: str
    qty: int
    due_date: str
    status: str
    proposed_at: str
    entered_at: str
    countersign_plan: str
    countersign_sales: str
    quality_approval: str
    released_by: str
    released_at: str


@dataclass(frozen=True)
class Schedule:
    schedule_id: str
    order_id: str
    planned_version: str
    qty: int
    line_id: str
    status: str
    updated_at: str


@dataclass(frozen=True)
class Capacity:
    line_id: str
    window_start: str
    window_end: str
    confirmed_capacity: int
    pending_overtime: int


@dataclass(frozen=True)
class Inventory:
    order_id: str
    sku: str
    batch_id: str
    available_qty: int
    wms_hold_status: str
    recorded_at: str


@dataclass(frozen=True)
class QualityHold:
    batch_id: str
    hold_status: str
    record_version: int
    effective_at: str
    recorded_at: str
    signed_by: str


@dataclass
class Snapshot:
    root: Path
    change_orders: list[ChangeOrder] = field(default_factory=list)
    schedules: list[Schedule] = field(default_factory=list)
    capacities: list[Capacity] = field(default_factory=list)
    inventories: list[Inventory] = field(default_factory=list)
    holds: list[QualityHold] = field(default_factory=list)
    file_hashes: dict[str, str] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)


def _int(value, what, errors):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        errors.append(f"{what} 不是整数：{value!r}")
        return None


def _enum(value, allowed, what, errors):
    value = (value or "").strip()
    if value not in allowed:
        errors.append(f"{what} 取值非法：{value!r}（允许 {sorted(allowed)}）")
        return None
    return value


def _req(row, key, what, errors):
    value = (row.get(key) or "").strip()
    if not value and key not in OPTIONAL_FIELDS:
        errors.append(f"{what} 缺少字段 {key}")
    return value


def _read(name, root, errors):
    path = root / name
    if not path.is_file():
        errors.append(f"缺少快照文件：{name}")
        return []
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            return [r for r in csv.DictReader(fh) if any((v or "").strip() for v in r.values())]
    except (OSError, UnicodeDecodeError, csv.Error) as exc:
        errors.append(f"快照文件无法读取：{name}（{exc}）")
        return []


def _build(cls, name, rows, specs, errors):
    out = []
    for i, row in enumerate(rows, start=1):
        errs = []
        values = {}
        for key, kind in specs:
            raw = _req(row, key, f"{name} 第 {i} 行", errs)
            if kind == OPT:  # 允许为空：未签发版本没有签发人/签发时间
                values[key] = raw
            elif kind == "int":
                values[key] = _int(raw, f"{name} 第 {i} 行 {key}", errs)
            elif isinstance(kind, set):
                values[key] = _enum(raw, kind, f"{name} 第 {i} 行 {key}", errs)
            else:
                values[key] = raw
        if errs:
            errors.extend(errs)
            out.append(None)  # 脏行占位：保留行号，规则层遇到即转人工
        else:
            out.append(cls(**values))
    return out


ERP_SPECS = [
    ("order_id", str), ("change_version", str), ("qty", "int"), ("due_date", str),
    ("status", ERP_STATUSES), ("proposed_at", str), ("entered_at", str),
    ("countersign_plan", YES_NO), ("countersign_sales", YES_NO), ("quality_approval", YES_NO),
    ("released_by", OPT), ("released_at", OPT),
]
MES_SPECS = [
    ("schedule_id", str), ("order_id", str), ("planned_version", str), ("qty", "int"),
    ("line_id", str), ("status", str), ("updated_at", str),
]
CAP_SPECS = [
    ("line_id", str), ("window_start", str), ("window_end", str),
    ("confirmed_capacity", "int"), ("pending_overtime", "int"),
]
WMS_SPECS = [
    ("order_id", str), ("sku", str), ("batch_id", str), ("available_qty", "int"),
    ("wms_hold_status", str), ("recorded_at", str),
]
QMS_SPECS = [
    ("batch_id", str), ("hold_status", HOLD_STATUSES), ("record_version", "int"),
    ("effective_at", str), ("recorded_at", str), ("signed_by", OPT),
]


def _hash(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_snapshot(root) -> Snapshot:
    root = Path(root)
    snap = Snapshot(root=root)
    for name in FILES:
        path = root / name
        if path.is_file():
            snap.file_hashes[name] = _hash(path)

    def rows(name):
        return _read(name, root, snap.errors)

    snap.change_orders = _build(ChangeOrder, "erp_change_orders.csv", rows(FILES[0]), ERP_SPECS, snap.errors)
    snap.schedules = _build(Schedule, "mes_schedule.csv", rows(FILES[1]), MES_SPECS, snap.errors)
    snap.capacities = _build(Capacity, "capacity.csv", rows(FILES[2]), CAP_SPECS, snap.errors)
    snap.inventories = _build(Inventory, "wms_inventory.csv", rows(FILES[3]), WMS_SPECS, snap.errors)
    snap.holds = _build(QualityHold, "qms_holds.csv", rows(FILES[4]), QMS_SPECS, snap.errors)
    return snap


def valid(items):
    return [x for x in items if x is not None]
