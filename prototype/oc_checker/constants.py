"""订单变更影响核对助手（只读原型）。

输入：ERP / MES / 产能 / WMS / QMS 五张脱敏快照（CSV）。
输出：核对报告（Markdown + JSON）与审计清单（manifest）。

约束（对应案例事实编号）：
- 只读，不回写任何生产系统（case-huaheng-f30 / f35）
- 模型/规则版本、输入快照、输出绑定同一任务编号（case-huaheng-f31）
- 输入缺失或非法时转人工，不给出部分结论冒充核对成功（case-huaheng-f32）
"""

RULES_VERSION = "rules-huaheng-1.0.0"

# 结论级别，严重程度由低到高
OK = "OK"
ATTENTION = "ATTENTION"
CONFLICT = "CONFLICT"
HARD_STOP = "HARD_STOP"
MANUAL_REVIEW = "MANUAL_REVIEW"
LEVELS = (OK, ATTENTION, CONFLICT, HARD_STOP, MANUAL_REVIEW)
LEVEL_LABEL = {
    OK: "一致",
    ATTENTION: "需关注",
    CONFLICT: "冲突",
    HARD_STOP: "硬约束中止",
    MANUAL_REVIEW: "转人工",
}

# 每类判断引用的访谈事实编号
FACTS = {
    "three_states": "case-huaheng-f14",
    "verbal_not_approved": "case-huaheng-f11",
    "version_conflict": "case-huaheng-f20",
    "capacity_gap": "case-huaheng-f21",
    "overtime_approval": "case-huaheng-f22",
    "countersign": "case-huaheng-f23",
    "plan_not_replaced": "case-huaheng-f25",
    "hold_hard": "case-huaheng-f40",
    "hold_sync": "case-huaheng-f41",
    "rework_not_pass": "case-huaheng-f42",
    "one_wrong_release": "case-huaheng-f43",
    "quality_scope": "case-huaheng-f44",
    "model_no_release": "case-huaheng-f45",
    "manual_on_failure": "case-huaheng-f32",
    "draft_not_sent": "case-huaheng-f13",
    "draft_mark_basis": "case-huaheng-f15",
    "confirmed_vs_pending": "case-huaheng-f52",
    "who_answers": "case-huaheng-f54",
}

# 人工关卡：工具只提示“待谁批”
OWNERS = {
    "countersign": "生产计划 + 销售（会签）",
    "quality_release": "质量负责人",
    "overtime": "工厂经理 + 人事",
    "gap_decision": "销售（向客户答复后回填）",
    "version_ruling": "会签双方（生产计划 + 销售）",
}
