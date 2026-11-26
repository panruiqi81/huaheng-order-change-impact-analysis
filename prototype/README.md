# 订单变更影响核对助手 · 只读原型

FDE 九阶段训练 · 华衡智造案例 · 阶段 05–06 产物。读取四系统脱敏快照，核对排产版本、产能缺口与质量冻结，输出“已确认事项 + 待审批事项（待谁批）”。**全程只读，不回写、不发送、不生成放行批准。**

> 仓库内数据均为按访谈事实构造的合成数据。

## 运行

需要 Python 3.10+，无第三方依赖。

```bash
cd prototype
python3 -m oc_checker --snapshot data/base --order O-202610-001 --out out/
```

输出 `out/report.md`（给人看）、`out/report.json`（给系统看）、`out/manifest.json`（审计清单）。

## 测试与评估

```bash
python3 -m unittest discover -s tests -v   # 7 项单元测试，覆盖 17 个评估用例
python3 tools/run_eval.py                  # 生成 eval/results.md 与 eval/base-report.md
```

评估结果见 [`eval/results.md`](eval/results.md)，基准场景报告见 [`eval/base-report.md`](eval/base-report.md)。

## 目录

```
oc_checker/     核对程序
  models.py     快照加载与校验
  rules.py      核对规则（版本 / 产能 / 质量 / 待确认清单）
  audit.py      任务编号与哈希绑定
  render.py     报告渲染
data/base/      合成基准快照（五张 CSV）
eval/           评估集（cases.py）、判定（verify.py）、结果
tests/          单元测试
tools/run_eval.py  评估脚本
```

## 快照字段

| 文件 | 字段 |
|---|---|
| `erp_change_orders.csv` | order_id, change_version, qty, due_date, status(proposed/countersigned/released), proposed_at, entered_at, countersign_plan(yes/no), countersign_sales(yes/no), quality_approval(yes/no), released_by, released_at |
| `mes_schedule.csv` | schedule_id, order_id, planned_version, qty, line_id, status, updated_at |
| `capacity.csv` | line_id, window_start, window_end, confirmed_capacity, pending_overtime |
| `wms_inventory.csv` | order_id, sku, batch_id, available_qty, wms_hold_status, recorded_at |
| `qms_holds.csv` | batch_id, hold_status(held/rework_pending_inspection/released), record_version, effective_at, recorded_at, signed_by |

## 结论级别

| 级别 | 含义 |
|---|---|
| OK | 一致，未发现问题 |
| ATTENTION | 需关注：有待批准变更、产能缺口、待批加班、状态同步延迟 |
| CONFLICT | 冲突：MES 排产版本与批准版本不一致，或订单没有批准版本 |
| HARD_STOP | 硬约束中止：订单占用批次处于冻结、返工未检或放行无签发人 |
| MANUAL_REVIEW | 转人工：输入不完整或不合法，本次核对未执行 |

每条发现都带对应的访谈事实编号（`case-huaheng-fXX`），见 `oc_checker/constants.py`。
