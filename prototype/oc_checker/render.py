"""报告渲染：给生产计划与客服看的 Markdown，以及给审计用的 JSON。"""

from __future__ import annotations

import json

from .constants import LEVEL_LABEL, MANUAL_REVIEW

STATUS_LABEL = {"proposed": "客户已提出", "countersigned": "会签完成·未签发", "released": "内部已批准"}

DISCLAIMER = (
    "> 本报告是只读核对结果，**不构成对外交期承诺**。对外答复须区分“已确认”与“待审批”，"
    "预计日期不能写成承诺日期；草稿由销售确认后才能发送。"
)


def _findings_table(findings):
    if not findings:
        return "_无。_\n"
    lines = ["| 结论 | 说明 | 依据 |", "|---|---|---|"]
    for f in findings:
        lines.append(f"| {LEVEL_LABEL[f['level']]} | {f['text']} | {'、'.join(f['facts']) or '—'} |")
    return "\n".join(lines) + "\n"


def to_markdown(result):
    task = result["task"]
    status = result["overall_status"]
    out = [f"# 订单变更影响核对报告 · {task['order_id']}", ""]

    def h(title):  # 标题自带节号；没有内容的章节不输出，后面的节号顺延
        out.extend([f"## {title}", ""])

    out += [
        "| 项 | 值 |", "|---|---|",
        f"| 任务编号 | `{task['task_id']}` |",
        f"| 规则版本 | `{task['rules_version']}` |",
        f"| 生成时间 | {task['generated_at']} |",
        f"| **结论** | **{LEVEL_LABEL[status]}** |",
        "",
        DISCLAIMER,
        "",
    ]

    h("一、已确认事项")
    out += [f"- {c}" for c in result["confirmed"]] or ["_本次没有可确认事项。_"]
    out.append("")

    if status == MANUAL_REVIEW and not result.get("version_check"):
        h("二、转人工原因")
        out += [f"- {r}" for r in result["manual_review_reasons"]]
        out += ["", "输入问题明细：", ""] + [f"- {e}" for e in result["input_errors"]]
        h("三、审计：输入快照")
        out += ["| 文件 | SHA-256 |", "|---|---|"]
        out += [f"| {k} | `{v}` |" for k, v in task["input_files"].items()]
        return "\n".join(out) + "\n"

    v, cap = result["version_check"], result["capacity_check"]
    h("二、版本一致性")
    out += [
        f"- ERP 已批准版本：**{v['approved_version'] or '无'}**；最新版本：{v['latest_version']}（{STATUS_LABEL.get(v['latest_status'], v['latest_status'])}）",
        "",
        _findings_table(v["findings"]),
    ]
    h("三、产能与缺口（按版本分别测算）")
    out += [
        "| 版本 | 状态 | 需求量 | 已确认产能 | 可用库存 | 合计 | 缺口 |",
        "|---|---|---|---|---|---|---|",
    ]
    for s in cap["scenarios"]:
        supply = s.get("supply")
        out.append(
            f"| {s['version']} | {STATUS_LABEL[s['status']]} | {s['qty']} | {cap['confirmed_capacity'] if cap['confirmed_capacity'] is not None else '未知'} "
            f"| {cap['usable_inventory']} | {supply if supply is not None else '未知'} | {s['gap'] if s['gap'] is not None else '未知'} |"
        )
    out += [
        "",
        "说明：未批准版本的数量只作情景测算，不构成产能承诺；待批加班不计入“已确认产能”。",
        "",
        _findings_table(cap["findings"]),
    ]
    h("四、质量冻结检查")
    out += [_findings_table(result["quality_findings"])]
    h("五、待审批事项（工具只提示“待谁批”，不替代审批）")
    if result["pending"]:
        out += ["| 编号 | 事项 | 答复人 | 说明 | 依据 |", "|---|---|---|---|---|"]
        for p in result["pending"]:
            out.append(f"| {p['id']} | {p['item']} | {p['owner']} | {p['detail']} | {'、'.join(p['facts'])} |")
    else:
        out.append("_无待审批事项。_")
    out.append("")

    has_manual = bool(result["manual_review_reasons"])
    if has_manual:
        h("六、需人工确认")
        out += [f"- {r}" for r in result["manual_review_reasons"]] + [""]

    h(("七" if has_manual else "六") + "、审计：输入快照")
    out += ["| 文件 | SHA-256 |", "|---|---|"]
    out += [f"| {k} | `{v}` |" for k, v in task["input_files"].items()]
    return "\n".join(out) + "\n"


def to_json(result):
    return json.dumps(result, ensure_ascii=False, indent=2)
