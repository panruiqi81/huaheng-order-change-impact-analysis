"""核对规则：版本一致性、产能缺口、质量冻结、待确认清单。"""

from __future__ import annotations

import re
from collections import defaultdict

from .constants import ATTENTION, CONFLICT, FACTS, HARD_STOP, LEVELS, MANUAL_REVIEW, OK, OWNERS
from .models import valid

STATUS_LABEL = {"proposed": "客户已提出", "countersigned": "会签完成·未签发", "released": "内部已批准"}


def worst(levels):
    ranks = [LEVELS.index(x) for x in levels if x in LEVELS]
    return LEVELS[max(ranks)] if ranks else OK


def version_no(version):
    m = re.search(r"(\d+)", version or "")
    return int(m.group(1)) if m else None


def finding(code, level, text, facts=()):
    return {"code": code, "level": level, "text": text, "facts": [FACTS[f] for f in facts]}


def check_version(orders, schedules):
    orders = valid(orders)
    schedules = valid(schedules)
    out = {"approved_version": None, "latest_version": None, "latest_status": None, "findings": []}
    if not orders:
        out["findings"].append(finding("NO_CHANGE_ORDER", MANUAL_REVIEW, "ERP 快照中没有该订单的变更单。", ("version_conflict",)))
        return out

    released = [o for o in orders if o.status == "released"]
    approved = max(released, key=lambda o: version_no(o.change_version) or -1) if released else None
    latest = max(orders, key=lambda o: version_no(o.change_version) or -1)
    out["approved_version"] = approved.change_version if approved else None
    out["latest_version"] = latest.change_version
    out["latest_status"] = latest.status

    f = out["findings"]
    if approved is None:
        f.append(finding("NO_APPROVED_VERSION", CONFLICT,
                         f"订单没有内部已批准版本：最新版本 {latest.change_version} 状态为“{STATUS_LABEL[latest.status]}”，"
                         "口头或群内提出的变更不能当作已批准版本。", ("three_states", "verbal_not_approved")))
    elif latest is not approved:
        note = "会签已完成但尚未签发" if latest.status == "countersigned" else "尚未会签"
        f.append(finding("PENDING_CHANGE", ATTENTION,
                         f"ERP 存在待批准的新版本 {latest.change_version}（{latest.qty} 件，{note}），"
                         f"当前批准版本仍为 {approved.change_version}（{approved.qty} 件）。", ("three_states", "countersign")))

    if not schedules:
        f.append(finding("NO_SCHEDULE", MANUAL_REVIEW, "MES 快照中没有该订单的排产记录，无法核对排产版本。"))
        return out

    for s in schedules:
        planned, base = version_no(s.planned_version), version_no(approved.change_version) if approved else None
        if approved is None:
            f.append(finding("SCHEDULE_WITHOUT_APPROVAL", CONFLICT,
                             f"排产 {s.schedule_id} 按 {s.planned_version}（{s.qty} 件）执行，但订单没有已批准版本。",
                             ("version_conflict", "plan_not_replaced")))
        elif planned is None or planned < base:
            f.append(finding("SCHEDULE_BEHIND", CONFLICT,
                             f"排产 {s.schedule_id} 仍按 {s.planned_version}（{s.qty} 件）执行，落后于已批准版本 "
                             f"{approved.change_version}（{approved.qty} 件）：现场在做旧版。", ("version_conflict",)))
        elif planned > base:
            f.append(finding("SCHEDULE_AHEAD", CONFLICT,
                             f"排产 {s.schedule_id} 已按 {s.planned_version}（{s.qty} 件）排产，但该版本尚未批准："
                             "建议方案不能自动替代批准的生产计划。", ("version_conflict", "plan_not_replaced")))
        else:
            f.append(finding("SCHEDULE_ALIGNED", OK,
                             f"排产 {s.schedule_id} 按 {s.planned_version}（{s.qty} 件）执行，与已批准版本一致。",
                             ("version_conflict",)))
    return out


def check_quality(inventories, holds):
    """返回 (findings, usable_batches, blocked_qty_by_batch)。冻结与返工未检批次不计入可用量。"""
    by_batch = defaultdict(list)
    for h in valid(holds):
        by_batch[h.batch_id].append(h)

    findings, usable, blocked = [], [], {}
    for inv in valid(inventories):
        recs = sorted(by_batch.get(inv.batch_id, []), key=lambda h: (h.record_version, h.effective_at))
        if not recs:
            findings.append(finding("NO_QUALITY_RECORD", MANUAL_REVIEW,
                                    f"批次 {inv.batch_id}（{inv.available_qty} 件）在 QMS 中没有质量记录，可用量待人工确认。",
                                    ("quality_scope",)))
            blocked[inv.batch_id] = inv.available_qty
            continue
        q = recs[-1]
        if q.hold_status == "held":
            findings.append(finding("BATCH_HELD", HARD_STOP,
                                    f"批次 {inv.batch_id} 处于质量冻结（记录 v{q.record_version}，{q.effective_at}），"
                                    f"其 {inv.available_qty} 件不得计入可用量；交期压力不能替代放行记录。",
                                    ("hold_hard", "one_wrong_release")))
            blocked[inv.batch_id] = inv.available_qty
        elif q.hold_status == "rework_pending_inspection":
            findings.append(finding("REWORK_NOT_INSPECTED", HARD_STOP,
                                    f"批次 {inv.batch_id} 返工完成但尚未重新检验（记录 v{q.record_version}），"
                                    f"其 {inv.available_qty} 件不得计入可用量：返工完成不等于检验通过。",
                                    ("rework_not_pass",)))
            blocked[inv.batch_id] = inv.available_qty
        elif q.hold_status == "released":
            if not q.signed_by:
                findings.append(finding("RELEASE_UNSIGNED", HARD_STOP,
                                        f"批次 {inv.batch_id} 的放行记录（v{q.record_version}）没有签发人，不视为有效放行；"
                                        f"其 {inv.available_qty} 件不计入可用量。", ("hold_hard", "model_no_release")))
                blocked[inv.batch_id] = inv.available_qty
            else:
                usable.append(inv)
        wms_held = inv.wms_hold_status == "held"
        qms_blocked = inv.batch_id in blocked
        if wms_held != qms_blocked:
            findings.append(finding("QMS_WMS_MISMATCH", ATTENTION,
                                    f"批次 {inv.batch_id} 的冻结状态在 QMS 与 WMS 不一致（QMS：{q.hold_status}，"
                                    f"WMS：{inv.wms_hold_status}），按同步延迟处理，以 QMS 带版本和时间的记录为准。",
                                    ("hold_sync",)))
    return findings, usable, blocked


def check_capacity(orders, schedules, capacities, usable_inventory, inventory_known=True):
    orders = valid(orders)
    caps = valid(capacities)
    lines = {s.line_id for s in valid(schedules)}
    out = {"findings": [], "scenarios": [], "confirmed_capacity": None, "pending_overtime": None, "usable_inventory": None}

    cap_rows = [c for c in caps if c.line_id in lines] if lines else []
    if not cap_rows:
        out["findings"].append(finding("NO_CAPACITY", MANUAL_REVIEW, "产能快照中没有该订单排产产线的产能记录，无法核算缺口。",
                                       ("capacity_gap",)))
    else:
        out["confirmed_capacity"] = sum(c.confirmed_capacity for c in cap_rows)
        out["pending_overtime"] = sum(c.pending_overtime for c in cap_rows)
        if out["pending_overtime"]:
            out["findings"].append(finding("OVERTIME_PENDING", ATTENTION,
                                           f"待批加班 {out['pending_overtime']} 件未计入可用产能：加班须经工厂经理与人事确认。",
                                           ("overtime_approval",)))
    out["usable_inventory"] = sum(i.available_qty for i in usable_inventory)
    if not inventory_known:
        # 缺记录不等于零库存：按 0 测算，但必须转人工确认
        out["findings"].insert(0, finding("NO_INVENTORY", MANUAL_REVIEW,
                                          "WMS 快照中没有该订单的库存记录，可用库存按 0 测算，须人工确认后才能使用本结果。",
                                          ("capacity_gap",)))
    supply_known = out["confirmed_capacity"] is not None

    # 对 ERP 中该订单的每个版本分别测算；未批准版本只是情景，不构成产能承诺
    for o in sorted(orders, key=lambda x: version_no(x.change_version) or -1):
        sc = {"version": o.change_version, "status": o.status, "qty": o.qty, "due_date": o.due_date, "level": OK, "gap": None}
        if supply_known:
            sc["supply"] = out["confirmed_capacity"] + out["usable_inventory"]
            sc["gap"] = max(0, o.qty - sc["supply"])
            if sc["gap"] > 0:
                sc["level"] = ATTENTION
                out["findings"].append(finding("CAPACITY_GAP", ATTENTION,
                                               f"按版本 {o.change_version}（{o.qty} 件）测算：已确认产能 {out['confirmed_capacity']} 件 "
                                               f"+ 可用库存 {out['usable_inventory']} 件 = {sc['supply']} 件，缺口 {sc['gap']} 件。",
                                               ("capacity_gap", "overtime_approval")))
        out["scenarios"].append(sc)
    return out


def build_lists(version, capacity, quality_findings, blocked):
    """已确认事项与待审批事项。对外答复须区分两者，且每个待确认项可查到答复人。"""
    confirmed, pending, seen = [], [], set()

    def add(key, item, owner, facts, detail=""):
        if key not in seen:
            seen.add(key)
            pending.append({"id": f"P{len(seen)}", "item": item, "owner": OWNERS[owner], "detail": detail,
                            "facts": [FACTS[f] for f in facts]})

    for f in version["findings"]:
        if f["code"] == "NO_APPROVED_VERSION":
            add("cs", f"变更单 {version['latest_version']} 的会签与签发", "countersign", ("countersign",),
                "会签完成前对外只能答复“已收到、正在确认”，不得按新版本承诺交期。")
        elif f["code"] == "PENDING_CHANGE":
            add("cs", f"变更单 {version['latest_version']} 的会签与签发", "countersign", ("countersign",),
                "批准版本由谁签发、以哪个系统为准尚未落定，需会签双方书面确认。")
        elif f["code"] in ("SCHEDULE_BEHIND", "SCHEDULE_AHEAD", "SCHEDULE_WITHOUT_APPROVAL"):
            add("vr", f"排产 {f['text'].split(' ')[1]} 与批准版本的裁定", "version_ruling", ("version_conflict", "plan_not_replaced"),
                "工具只列出冲突，是否调整排产由会签双方决定。")
        elif f["code"] == "SCHEDULE_ALIGNED":
            confirmed.append(f["text"])
    if version["approved_version"]:
        confirmed.append(f"ERP 当前已批准版本为 {version['approved_version']}。")

    gaps = [s for s in capacity["scenarios"] if s.get("gap")]
    if gaps:
        latest = gaps[-1]
        confirmed.append(f"缺口 {latest['gap']} 件（按版本 {latest['version']} 测算）未被默认补齐，显式列出。")
        add("gap", f"缺口 {latest['gap']} 件的处置：让交期还是让数量", "gap_decision", ("capacity_gap",),
            "由销售向客户答复后回填；工具不替客户决定。")
        if capacity["pending_overtime"]:
            add("ot", f"加班 {capacity['pending_overtime']} 件能否计入可用产能", "overtime", ("overtime_approval",),
                "未获批前不得计入可用产能。")
    for f in quality_findings:
        if f["level"] == HARD_STOP:
            item = "冻结批次的放行" if f["code"] == "BATCH_HELD" else "批次质量记录的补正"
            add("q", item, "quality_release",
                ("hold_hard", "one_wrong_release"), "工具只解释已有质量记录，不生成放行批准。")
    return confirmed, pending


def evaluate(snap, order_id):
    task = {"order_id": order_id}
    result = {"task": task, "overall_status": OK, "input_errors": list(snap.errors),
              "manual_review_reasons": [], "confirmed": [], "pending": []}

    orders = [o for o in valid(snap.change_orders) if o.order_id == order_id]
    schedules = [s for s in valid(snap.schedules) if s.order_id == order_id]
    inventories = [i for i in valid(snap.inventories) if i.order_id == order_id]

    dirty = any(x is None for x in (*snap.change_orders, *snap.schedules, *snap.capacities, *snap.inventories, *snap.holds))
    if snap.errors or dirty:
        result["input_errors"] = list(snap.errors)
        if dirty:
            result["input_errors"].append("快照中存在无法解析的行（已在上文逐行列出）。")
        result["overall_status"] = MANUAL_REVIEW
        result["manual_review_reasons"] = [
            "输入快照不完整或不合法，本次核对未执行，转人工处理。",
            "未使用任何缓存或历史结果充当本次核对结论。",
        ]
        result["version_check"], result["capacity_check"] = None, None
        result["quality_findings"] = []
        return result

    version = check_version(orders, schedules)
    quality_findings, usable, blocked = check_quality(inventories, snap.holds)
    capacity = check_capacity(orders, schedules, snap.capacities, usable, inventory_known=bool(inventories))
    confirmed, pending = build_lists(version, capacity, quality_findings, blocked)

    levels = [f["level"] for f in version["findings"]]
    levels += [f["level"] for f in quality_findings] + [f["level"] for f in capacity["findings"]]
    result["overall_status"] = worst(levels)
    result["version_check"] = version
    result["quality_findings"] = quality_findings
    result["capacity_check"] = capacity
    result["confirmed"] = confirmed
    result["pending"] = pending
    result["blocked_qty_by_batch"] = blocked
    result["manual_review_reasons"] = [
        f["text"] for f in (*version["findings"], *quality_findings, *capacity["findings"]) if f["level"] == MANUAL_REVIEW
    ]
    return result
