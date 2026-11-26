# 评估结果 · 订单变更影响核对助手（只读原型）

- 运行时间戳：2026-11-26T04:00:00+00:00（评估固定值，便于复现）
- 规则版本：`rules-huaheng-1.0.0` · Python 3.12.11 · Linux 6.6.95.bck.4-rc5-amd64
- 用例数：17 · 通过：**17** · 失败：0
- 结论分布：ATTENTION 3、CONFLICT 3、HARD_STOP 4、MANUAL_REVIEW 6、OK 1
- 路径分布：异常路径 A · 版本冲突 2、异常路径 B · 未批先排 1、异常路径 C · 缺口依赖加班 1、异常路径 D · 质量冻结 3、异常路径 E · 冻结状态同步延迟 1、异常路径 F · 数据缺失转人工 6、正常路径 3

| 用例 | 场景 | 路径 | 预期 | 实际 | 实际命中的发现编号 | 结果 |
|---|---|---|---|---|---|---|
| TC-01 | 基准：V3 已提出未会签，MES 按 V2 排产，缺口 100 件 | 正常路径 | ATTENTION | ATTENTION | PENDING_CHANGE, SCHEDULE_ALIGNED, OVERTIME_PENDING, CAPACITY_GAP | ✅ 通过 |
| TC-02 | V3 已批准，MES 仍按 V2 排产（现场在做旧版） | 版本冲突 | CONFLICT | CONFLICT | SCHEDULE_BEHIND, OVERTIME_PENDING, CAPACITY_GAP | ✅ 通过 |
| TC-03 | V3 未批准，MES 已按 V3 排产（超前排产） | 未批先排 | CONFLICT | CONFLICT | PENDING_CHANGE, SCHEDULE_AHEAD, OVERTIME_PENDING, CAPACITY_GAP | ✅ 通过 |
| TC-04 | 订单没有任何已批准版本 | 版本冲突 | CONFLICT | CONFLICT | NO_APPROVED_VERSION, SCHEDULE_WITHOUT_APPROVAL, OVERTIME_PENDING, CAPACITY_GAP | ✅ 通过 |
| TC-05 | V3 会签完成但尚未签发，不算批准 | 正常路径 | ATTENTION | ATTENTION | PENDING_CHANGE, SCHEDULE_ALIGNED, OVERTIME_PENDING, CAPACITY_GAP | ✅ 通过 |
| TC-06 | V3 已批准且 MES 已切换，产能充足 | 正常路径 | OK | OK | SCHEDULE_ALIGNED | ✅ 通过 |
| TC-07 | 待批加班 300 件足以覆盖缺口，但未获批不得计入产能 | 缺口依赖加班 | ATTENTION | ATTENTION | PENDING_CHANGE, SCHEDULE_ALIGNED, OVERTIME_PENDING, CAPACITY_GAP | ✅ 通过 |
| TC-08 | 订单占用冻结批次 B-17（QMS 与 WMS 一致） | 质量冻结 | HARD_STOP | HARD_STOP | PENDING_CHANGE, SCHEDULE_ALIGNED, BATCH_HELD, OVERTIME_PENDING, CAPACITY_GAP | ✅ 通过 |
| TC-09 | B-17 冻结，但 WMS 仍显示已放行（同步延迟） | 冻结状态同步延迟 | HARD_STOP | HARD_STOP | PENDING_CHANGE, SCHEDULE_ALIGNED, BATCH_HELD, QMS_WMS_MISMATCH, OVERTIME_PENDING, CAPACITY_GAP | ✅ 通过 |
| TC-10 | 批次返工完成但未重新检验 | 质量冻结 | HARD_STOP | HARD_STOP | PENDING_CHANGE, SCHEDULE_ALIGNED, REWORK_NOT_INSPECTED, QMS_WMS_MISMATCH, OVERTIME_PENDING, CAPACITY_GAP, CAPACITY_GAP | ✅ 通过 |
| TC-11 | 放行记录没有签发人，不视为有效放行 | 质量冻结 | HARD_STOP | HARD_STOP | PENDING_CHANGE, SCHEDULE_ALIGNED, RELEASE_UNSIGNED, QMS_WMS_MISMATCH, OVERTIME_PENDING, CAPACITY_GAP, CAPACITY_GAP | ✅ 通过 |
| TC-12 | 批次在 QMS 中没有质量记录 | 数据缺失转人工 | MANUAL_REVIEW | MANUAL_REVIEW | PENDING_CHANGE, SCHEDULE_ALIGNED, NO_QUALITY_RECORD, OVERTIME_PENDING, CAPACITY_GAP, CAPACITY_GAP | ✅ 通过 |
| TC-13 | 订单在 WMS 没有库存记录（缺记录 ≠ 零库存） | 数据缺失转人工 | MANUAL_REVIEW | MANUAL_REVIEW | PENDING_CHANGE, SCHEDULE_ALIGNED, NO_INVENTORY, OVERTIME_PENDING, CAPACITY_GAP, CAPACITY_GAP | ✅ 通过 |
| TC-14 | 缺少产能快照文件 | 数据缺失转人工 | MANUAL_REVIEW | MANUAL_REVIEW | — | ✅ 通过 |
| TC-15 | 变更单数量字段不是整数（脏数据） | 数据缺失转人工 | MANUAL_REVIEW | MANUAL_REVIEW | — | ✅ 通过 |
| TC-16 | MES 没有该订单的排产记录 | 数据缺失转人工 | MANUAL_REVIEW | MANUAL_REVIEW | PENDING_CHANGE, NO_SCHEDULE, NO_CAPACITY | ✅ 通过 |
| TC-17 | 查询不存在的订单号 | 数据缺失转人工 | MANUAL_REVIEW | MANUAL_REVIEW | NO_CHANGE_ORDER, NO_INVENTORY, NO_CAPACITY | ✅ 通过 |

## 评估口径

1. 每个用例由基准快照（`data/base/`）派生，改动项写在 `eval/cases.py`，可逐条复现。
2. 通过条件：总体结论、必须出现/不得出现的发现编号、关键字段值全部符合预期，且运行前后输入快照哈希不变（只读）。
3. 预期值是先写后测的：先按访谈事实定下每个用例“应该得出什么”，再运行比对；失败项原样保留，不改预期去迁就结果。
4. 评估对象是规则引擎的确定性输出，不涉及大模型；接入模型生成说明文字后需另建评估集。
