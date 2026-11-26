# 生成《企业报价包》全部 CSV，并自检数字咬合。
# 用法：pwsh -NoProfile -ExecutionPolicy Bypass -File pricing/tools/build_pricing.ps1
param(
    [string]$OutDir = (Split-Path -Parent $PSScriptRoot)
)
$ErrorActionPreference = 'Stop'
$inv = [System.Globalization.CultureInfo]::InvariantCulture

function Q([object]$s) { '"' + ([string]$s -replace '"', '""') + '"' }
function R0([double]$x) { [int][math]::Round($x, 0, [MidpointRounding]::AwayFromZero) }
function R1([double]$x) { [math]::Round($x, 1, [MidpointRounding]::AwayFromZero).ToString('0.0', $inv) }
function R2([double]$x) { [math]::Round($x, 2, [MidpointRounding]::AwayFromZero).ToString('0.00', $inv) }
function N0([double]$x) { $x.ToString('N0', $inv) }
function N1([double]$x) { $x.ToString('N1', $inv) }
function R4([double]$x) { $x.ToString('0.0000', $inv) }
function Write-Csv([string]$name, [object[]]$rows) {
    $lines = foreach ($r in $rows) { (($r | ForEach-Object { Q $_ }) -join ',') }
    $path = Join-Path $OutDir $name
    [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), [System.Text.UTF8Encoding]::new($true))
    $Tables[$name] = $rows
    Write-Host ("已生成 " + $name + "（" + $rows.Count + " 行）")
}
function New-Rows { return ,([System.Collections.Generic.List[object]]::new()) }
$Tables = @{}   # 各表的行数据，供生成 Excel 工作簿与 Markdown 版复用

# ---------------------------------------------------------------- 口径参数（教学假设，见 README 假设清单）
$coef = 1.55        # 综合成本系数
$util = 0.80        # 可计费利用率
$dpm = 21.75        # 每月可计费工作日
$markup = 0.25      # 成本加成率
$risk = 0.15        # 风险预留率
$bufRate = 0.20     # 整体缓冲

# ---------------------------------------------------------------- 角色单价推导
$ref = [ordered]@{
    '解决方案架构师'   = @(30000, '2,500–3,500')
    'FDE 工程师'      = @(22000, '1,800–2,500')
    '产品经理'        = @(18000, '1,800–2,500')
    '后端工程师'      = @(16000, '1,500–2,200')
    '测试/实施工程师' = @(12000, '1,200–1,800')
}
$cost = @{}; $price = @{}; $monthCost = @{}
foreach ($k in $ref.Keys) {
    $monthCost[$k] = $ref[$k][0] * $coef
    $cost[$k] = R0 ($monthCost[$k] / $dpm / $util)
    $price[$k] = R0 ($cost[$k] * (1 + $markup))
}

# ---------------------------------------------------------------- 工作包（可计数）
$tasks = @(
    @{ g = '需求调研与方案设计'; n = '业务调研与角色访谈（6 个角色 × 1 轮）'; r = 'FDE 工程师'; d = 4; note = '总经理、销售、生产计划、信息化、质量、客服' }
    @{ g = '需求调研与方案设计'; n = '现状流程还原与四来源口径收敛'; r = '产品经理'; d = 2; note = '交付率两套口径先锁定' }
    @{ g = '需求调研与方案设计'; n = '需求规格、追溯矩阵与验收指标'; r = '产品经理'; d = 3; note = '含停止条件与试点范围' }
    @{ g = '需求调研与方案设计'; n = '方案设计与备选比较'; r = '解决方案架构师'; d = 2; note = '含安全边界与人工关卡设计' }
    @{ g = '知识库/规则库构建'; n = '核对规则实现 ×4（版本一致性、质量冻结、产能缺口、待确认清单）'; r = 'FDE 工程师'; d = 8; note = '确定性规则，核对结论不依赖模型' }
    @{ g = '核对流程开发（按流程条数拆）'; n = '报告说明生成 ×1（模型调用、提示词、输出约束、回退模板）'; r = 'FDE 工程师'; d = 3; note = '模型只解释已有结果，不参与判定' }
    @{ g = '系统对接（按接口个数拆）'; n = '四系统快照接入与字段口径对齐 ×4（ERP / MES / WMS / QMS）'; r = '后端工程师'; d = 6; note = '只读快照，不含写接口' }
    @{ g = '系统对接（按接口个数拆）'; n = '审计留痕与任务编号绑定'; r = '后端工程师'; d = 2; note = '每次核对可追溯' }
    @{ g = '系统对接（按接口个数拆）'; n = '报告输出（Markdown / JSON）与命令行入口'; r = '后端工程师'; d = 2; note = '' }
    @{ g = '集成联调（单独列）'; n = '真实快照字段口径联调'; r = '后端工程师'; d = 3; note = '约占开发量 20%–30%' }
    @{ g = '集成联调（单独列）'; n = '业务复核联调（与生产计划、客服逐单比对）'; r = 'FDE 工程师'; d = 2; note = '' }
    @{ g = '测试评估'; n = '评估集构建（17 个用例，先写预期）'; r = '测试/实施工程师'; d = 3; note = '预期结论先于运行写定' }
    @{ g = '测试评估'; n = '单元测试与红线测试（只读、审计、转人工）'; r = '测试/实施工程师'; d = 2; note = '' }
    @{ g = '测试评估'; n = '红线评审与质量复核'; r = '解决方案架构师'; d = 1; note = '质量负责人参与' }
    @{ g = '效果调优（含 2 轮）'; n = '规则口径调优（2 轮，每轮 1 天）'; r = 'FDE 工程师'; d = 2; note = '超出 2 轮按人天另计' }
    @{ g = '试点与交接'; n = '试点运行支持（4 周，每周运行与复盘）'; r = '测试/实施工程师'; d = 4; note = '' }
    @{ g = '试点与交接'; n = '培训与交接材料（运维手册、培训）'; r = 'FDE 工程师'; d = 2; note = '' }
    @{ g = '试点与交接'; n = '上线检查与放行门禁评审'; r = '解决方案架构师'; d = 1; note = '生产计划、质量、信息化三方' }
)
$byRole = [ordered]@{}
foreach ($k in $ref.Keys) { $byRole[$k] = @{ days = 0; cost = 0; price = 0 } }
$sumDays = 0; $laborCost = 0; $laborPrice = 0
$devDays = 0; $intDays = 0
foreach ($t in $tasks) {
    $c = $t.d * $cost[$t.r]; $p = $t.d * $price[$t.r]
    $byRole[$t.r].days += $t.d; $byRole[$t.r].cost += $c; $byRole[$t.r].price += $p
    $sumDays += $t.d; $laborCost += $c; $laborPrice += $p
    if ($t.g -match '规则库|流程开发|系统对接') { $devDays += $t.d }
    if ($t.g -match '集成联调') { $intDays += $t.d }
}
$wc = R0 ($laborCost / $sumDays); $wp = R0 ($laborPrice / $sumDays)
$bufDays = R0 ($sumDays * $bufRate)
$bufCost = $bufDays * $wc; $bufPrice = $bufDays * $wp
$bufRole = @{}; foreach ($k in $byRole.Keys) { $bufRole[$k] = $bufDays * $byRole[$k].days / $sumDays }
$laborCostTotal = $laborCost + $bufCost; $laborPriceTotal = $laborPrice + $bufPrice
$totalDays = $sumDays + $bufDays
$intShare = $intDays / $devDays

# ---------------------------------------------------------------- Token 与模型
$inParts = @(
    @{ n = '系统提示词（角色、约束、输出格式）'; t = 1200; s = '固定前缀，可命中缓存' }
    @{ n = '结构化核对结果（已确认 / 发现 / 待审批清单）'; t = 1500; s = '规则引擎输出，已脱敏' }
    @{ n = '规则说明与事实依据'; t = 800; s = '规则版本 1.0.0 的条文摘录' }
    @{ n = '订单号与本次指令'; t = 100; s = '' }
    @{ n = '历史对话'; t = 0; s = '单次调用，不携带历史' }
    @{ n = '工具返回结果'; t = 0; s = '已包含在结构化核对结果内' }
)
$inTok = 0; foreach ($p in $inParts) { $inTok += $p.t }
$outTok = 800
$pIn = 1.08; $pOut = 4.32          # DeepSeek Flash 离峰价（元 / 百万 Token）
$steps = 1; $retry = 0.05; $cacheHit = 0
$perCall = $inTok * $pIn / 1e6 + $outTok * $pOut / 1e6
$perReq = $perCall * $steps * (1 + $retry)
$ordersPerMonth = 40                # 月核对报告量（假设，与 ROI 一致）
$daily = $ordersPerMonth / 30       # 日均请求数
$levels = @(
    @{ name = '保守'; f = 0.7 }
    @{ name = '基准'; f = 1.0 }
    @{ name = '高峰'; f = 1.6 }
)
$tokRows = @{}
foreach ($l in $levels) {
    $req = $daily * 30 * $l.f
    $m = $perReq * $req
    $tokRows[$l.name] = @{ req = $req; month = $m; year = $m * 12; tok = $req * ($inTok + $outTok) }
}
$tokBaseMonth = $tokRows['基准'].month
$tokBaseYear = $tokRows['基准'].year
$tokBaseTokMonth = $tokRows['基准'].tok
$modelPack = $tokBaseYear * 1.20          # 用量包干含 20% 缓冲
$apiBlend = ($inTok * $pIn + $outTok * $pOut) / ($inTok + $outTok)   # 综合单价（元 / 百万 Token）

# ---------------------------------------------------------------- 服务器、数据、运维、杂项（成本口径）
$devEnv = 120 * 3                       # 开发测试云主机 3 个月
$prodEnv = 0                            # 生产环境用客户内网现有虚拟机
$dataCost = 0
$opsResp = 0.5 * 12 * $cost['FDE 工程师']
$opsIter = 4 * $cost['FDE 工程师']
$opsCost = $opsResp + $opsIter
$travel = 4 * 1500
$annualServiceCost = $opsCost + $modelPack
$firstYearCost = $laborCostTotal + $modelPack + $devEnv + $dataCost + $opsCost + $travel
$renewCost = $annualServiceCost
$implFee = R0 (($laborCostTotal + $devEnv + $travel) * (1 + $risk) * (1 + $markup))
$implCostBase = $laborCostTotal + $devEnv + $travel
$serviceFee = R0 ($annualServiceCost * (1 + $risk) * (1 + $markup))
$firstYearQuote = $implFee + $serviceFee

# ---------------------------------------------------------------- 部署方案
$gpuBuy = 25000; $depYears = 3; $powerKw = 0.6; $elecPrice = 0.8; $pue = 1.4
$elecYear = $powerKw * 24 * 365 * $elecPrice * $pue
$modelOps = 1000 * 12
$deployDaysB = 5; $deployCostB = $deployDaysB * $cost['FDE 工程师']
$weightGb = 7; $kvKb = 56; $ctx = 8192; $conc = 2
$kvGb = $kvKb * $ctx * $conc / 1e6
$vramGb = ($weightGb + $kvGb) * 1.15
$fixedYearB = $gpuBuy / $depYears + $elecYear + $modelOps
$threeYearB = $gpuBuy + 3 * ($elecYear + $modelOps) + $deployCostB
$fixedMonthB = $fixedYearB / 12
$critTok = $fixedMonthB / $apiBlend * 1e6
$mixShare = 0.30
$schemeAYear = $tokBaseYear
$schemeCYear = $tokBaseYear * $mixShare

# ---------------------------------------------------------------- ROI（客户侧）
$people = 6; $monthSalary = 8000
$annualCostPP = $monthSalary * 12 * 1.4
$hourCost = $annualCostPP / 2000             # 综合时薪（元 / 小时）
$hoursPerOrder = 2.5                         # 每单协调工时（变更前，销售 / 计划 / 客服 / 工厂合计，假设）
$hoursCut = 0.6                              # 上线后协调工时压降比例（假设）
$hoursSavedMonth = $ordersPerMonth * $hoursPerOrder * $hoursCut
$saveRatio = $hoursSavedMonth / ($people * 2000 / 12)
$fteSaved = $hoursSavedMonth / (2000 / 12)
$conv = 0.5
$gainLabor = $hoursSavedMonth * 12 * $hourCost * $conv
$errCount = 12; $errLoss = 20000; $errCut = 0.5
$gainErr = $errCount * $errLoss * $errCut
$benefit = $gainLabor + $gainErr
$I1 = $implFee; $runCost = $serviceFee; $firstTotal = $I1 + $runCost
function Get-Roi([double]$b) {
    @{
        benefit = $b
        roi1 = ($b - $firstTotal) / $firstTotal
        payback = $firstTotal / ($b / 12)
        three = (3 * $b - ($I1 + 3 * $runCost)) / ($I1 + 3 * $runCost)
        ratio = $b / $runCost
        netMonth = ($b - $runCost) / 12
        paybackCurve = $(if ($b - $runCost -gt 0) { $I1 / (($b - $runCost) / 12) } else { [double]::NaN })
    }
}
$roiBase = Get-Roi $benefit
$roiLow = Get-Roi ($benefit * 0.7)
$roiHigh = Get-Roi ($benefit * 1.3)
$roiLaborOnly = Get-Roi $gainLabor
# 首年回本的前提：只靠工时、或只靠避免冲突，各需要多大的量
$beHours = $firstTotal / 12 / ($hourCost * $conv)
$beOrders = $beHours / ($hoursPerOrder * $hoursCut)
$beLoss = $firstTotal / ($errCount * $errCut)
$beCount = $firstTotal / ($errLoss * $errCut)
function Pct([double]$x) { ($x * 100).ToString('0.0', $inv) + '%' }
function Mon([double]$x) { if ([double]::IsNaN($x)) { '不回本' } else { $x.ToString('0.0', $inv) } }

# ================================================================ 00 项目一页纸
$rows = New-Rows
$rows.Add(@('要填的五项', '华衡智造 · 订单变更影响核对助手（本项目）', '填写依据 / 待确认'))
$rows.Add(@('客户是谁', '华衡智造（制造企业，规模未在案例中给出，按约 300 人中型制造企业假设）；总经理周明远出资，生产计划部牵头使用', '访谈：问题由销售总监提出，生产计划经理为主要用户；规模为假设'))
$rows.Add(@('多少人用', '约 7 人：生产计划 2、客户服务 3、销售 1、质量 1', '案例出现 6 个角色，各岗位人数为假设'))
$rows.Add(@('每天多少量', '客服日均约 45 条变更咨询；核对报告按月 40 份（约 1.3 份/日）', '45 条为客服主管陈述；月 40 单为假设，与阶段 04 的 N=40 一致'))
$rows.Add(@('数据在哪', 'ERP、MES、WMS、QMS 均在内网，仅提供脱敏 CSV 或只读快照，不提供写接口', '访谈事实：四系统只给脱敏快照'))
$rows.Add(@('能不能上云', '核心核对不上云（纯规则，内网运行）；仅“报告说明文字”脱敏后调云端 API，须信息化负责人确认，未确认则回退固定模板', '待确认项；见 04 部署方案对比表'))
Write-Csv '00-项目一页纸.csv' $rows

# ================================================================ 02 人天估算与单价表
$rows = New-Rows
$rows.Add(@('阶段（工作包类别）', '序号', '工作包（可计数）', '角色', '天数', '单价-成本口径(元)', '单价-对外口径(元)', '小计-成本口径(元)', '小计-对外口径(元)', '说明'))
$i = 0
foreach ($t in $tasks) {
    $i++
    $rows.Add(@($t.g, $i, $t.n, $t.r, $t.d, $cost[$t.r], $price[$t.r], ($t.d * $cost[$t.r]), ($t.d * $price[$t.r]), $t.note))
}
$rows.Add(@('小计', '', ($tasks.Count.ToString() + ' 个工作包'), '', $sumDays, $wc, $wp, $laborCost, $laborPrice, ('加权平均单价 = 小计 ÷ ' + $sumDays + ' 人天')))
$rows.Add(@('整体缓冲', 19, '整体缓冲 20%', '按加权单价', $bufDays, $wc, $wp, $bufCost, $bufPrice, '需求侧尚有 6 项缺口未关闭，取 15%–25% 区间中值；不摊入各工作包'))
$rows.Add(@('合计', '', '含缓冲', '', $totalDays, '', '', $laborCostTotal, $laborPriceTotal, '报价以成本口径为基数，风险预留与加成各乘一次（见 01 表）'))
Write-Csv '02-人天估算与单价表.csv' $rows

$rows = New-Rows
$rows.Add(@('角色', '月薪(元)', '综合成本系数', '可计费利用率(%)', '成本加成率(%)', '月综合成本(元)', '人天综合成本(元)', '对外人天单价(元)', '课程参考区间(元)'))
foreach ($k in $ref.Keys) {
    $rows.Add(@($k, $ref[$k][0], $coef.ToString('0.00', $inv), ($util * 100), ($markup * 100), (R0 $monthCost[$k]), $cost[$k], $price[$k], $ref[$k][1]))
}
$rows.Add(@('算法工程师', '不列', '', '', '', '', '', '', '2,000–3,000'))
$rows.Add(@('说明：人天综合成本 = 月薪 × 综合成本系数 ÷ 21.75 ÷ 利用率；对外人天单价 = 人天综合成本 × (1 + 成本加成率)。FDE 一行与测算器导出结果一致（1,960 / 2,450）。本项目不涉及模型微调，故不列算法工程师。'))
Write-Csv '02-人天单价推导.csv' $rows

# ================================================================ 01 成本拆解表
$rows = New-Rows
$rows.Add(@('成本拆解表', '明细项', '角色', '天数', '单价(元)', '测算依据', '金额(元)', '计费周期', '备注'))
$first = $true
foreach ($t in $tasks) {
    $basis = if ($first) { '人天综合成本 = 月薪 × 1.55 ÷ 21.75 ÷ 0.8（02 表）；' + $t.d + ' 天 × ' + $cost[$t.r] } else { '' + $t.d + ' 天 × ' + $cost[$t.r] }
    $rows.Add(@('人力', ($t.g + '：' + $t.n), $t.r, $t.d, $cost[$t.r], $basis, ($t.d * $cost[$t.r]), '首年（一次性）', $t.note))
    $first = $false
}
$rows.Add(@('人力', '整体缓冲 20%', '按加权单价', $bufDays, $wc, ('' + $sumDays + ' 人天 × 20% = ' + $bufDays + ' 人天'), $bufCost, '首年（一次性）', '含联调与调优，不摊入各工作包'))
$rows.Add(@('人力', '人力小计', '', $totalDays, '', ('工作包 ' + $sumDays + ' 人天 + 缓冲 ' + $bufDays + ' 人天（成本口径）'), $laborCostTotal, '首年（一次性）', ('联调 ' + $intDays + ' 人天，占开发量 ' + ($intShare * 100).ToString('0', $inv) + '%')))
$rows.Add(@('模型', '报告说明生成（DeepSeek Flash，离峰价）', '', '', '', ('03 表基准档 ' + (R2 $tokBaseMonth) + ' 元/月 × 12'), (R2 $tokBaseYear), '首年起每年', '核对结论不调用模型'))
$rows.Add(@('模型', '用量缓冲 20%', '', '', '', '包干量含 20% 用量缓冲，超量按单价结算', (R2 ($modelPack - $tokBaseYear)), '首年起每年', ''))
$rows.Add(@('服务器', '开发测试环境', '', '', '', '云主机 2 核 4G 约 120 元/月（假设）× 3 个月', $devEnv, '首年（一次性）', '待按云平台实时报价复核'))
$rows.Add(@('服务器', '生产运行环境', '', '', '', '核对程序仅用标准库 Python，部署在客户内网现有虚拟机（客户自备）', $prodEnv, '—', '无 GPU、无云租'))
$rows.Add(@('数据', '脱敏快照导出与字段口径对齐', '', '', '', '已含在人力“四系统快照接入 ×4”（6 人天），不重复计', 0, '—', '数据块逐项列示，避免漏算'))
$rows.Add(@('数据', '第三方接口（发票查验、行业数据等）', '', '', '', '本项目不涉及', 0, '—', ''))
$rows.Add(@('运维', '日常响应与巡检', 'FDE 工程师', 6, $cost['FDE 工程师'], ('0.5 人天/月 × 12 个月 × ' + $cost['FDE 工程师']), $opsResp, '首年起每年', ''))
$rows.Add(@('运维', '版本迭代', 'FDE 工程师', 4, $cost['FDE 工程师'], ('4 人天/年 × ' + $cost['FDE 工程师']), $opsIter, '首年起每年', ''))
$rows.Add(@('实施差旅', '现场调研、试点复盘', '', '', '', '4 次 × 1,500 元（假设）', $travel, '首年（一次性）', ''))
$rows.Add(@('合计', '首年成本合计（成本口径）', '', '', '', '人力 + 模型 + 服务器 + 数据 + 运维 + 差旅，逐项加总', (R2 $firstYearCost), '首年', ''))
$rows.Add(@('合计', '续年成本合计（成本口径）', '', '', '', '运维 + 模型', (R2 $renewCost), '续年每年', ''))
$rows.Add(@('对外报价', '一次性实施费', '', $totalDays, '', ('成本 ' + (N0 $implCostBase) + ' × (1 + 15%) × (1 + 25%)'), $implFee, '首年（一次性）', '对外人天单价已含 25% 加成，故以成本口径为基数只加成一次'))
$rows.Add(@('对外报价', '年服务费', '', '', '', ('成本 ' + (R2 $annualServiceCost) + ' × (1 + 15%) × (1 + 25%)'), $serviceFee, '首年及续年每年', '含运维 10 人天/年与模型用量包'))
$rows.Add(@('对外报价', '首年合计', '', '', '', '实施费 + 年服务费', $firstYearQuote, '首年', ('续年 = 年服务费 ' + $serviceFee)))
$rows.Add(@('报价模式', '组合报价：一次性实施费 + 年服务费 + 超量按单价', '', '', '', '', '', '', '四选一取“项目总价制 + 订阅制”的组合'))
$rows.Add(@('选择理由(一句话)', '范围与验收标准清晰（17 个评估用例、只读核对），适合固定价交付；模型与运维成本按年发生，用年服务费覆盖，用量超包另计。', '', '', '', '', '', '', ''))
Write-Csv '01-成本拆解表.csv' $rows

# ================================================================ 03 Token
$rows = New-Rows
$rows.Add(@('构成项', 'Token', '说明'))
foreach ($p in $inParts) { $rows.Add(@($p.n, $p.t, $p.s)) }
$rows.Add(@('单次输入合计', $inTok, '估算值，上线前用真实报告样本按分词器实测后重算'))
$rows.Add(@('单次输出', $outTok, '约 1,000 字中文说明'))
Write-Csv '03-单次输入构成.csv' $rows

$rows = New-Rows
$rows.Add(@('档位', '模型', '计价口径', '日均请求数(次)', '月请求数(次)', '单次输入Token', '单次输出Token', 'Agent步数', '重试率(%)', '输入缓存命中率(%)', '输入单价(元/百万Token)', '输出单价(元/百万Token)', '单次请求成本(元)', '月成本(元)', '年成本(元)', '月Token总量'))
foreach ($l in $levels) {
    $x = $tokRows[$l.name]
    $mult = if ($l.f -eq 1.0) { '业务量 ×1.0' } else { '业务量 ×' + $l.f.ToString('0.0', $inv) }
    $rows.Add(@($l.name, 'DeepSeek deepseek-flash', $mult, $daily, (R2 $x.req), $inTok, $outTok, $steps, ($retry * 100), ($cacheHit * 100), $pIn, $pOut, (R4 $perReq), (R2 $x.month), (R2 $x.year), (R0 $x.tok)))
}
$rows.Add(@('说明：单次请求成本 = (输入 × 输入单价 + 输出 × 输出单价) × Agent 步数 × (1 + 重试率)；月成本 = 单次请求成本 × 日均请求数 × 30；年成本 = 月成本 × 12。基准档年成本即 06 报价汇总表里的模型费。价格为课件口径（2026-11-26 核验）的离峰价，高峰时段 ×2。'))
Write-Csv '03-Token月成本测算表.csv' $rows

# ================================================================ 04 部署方案
$rows = New-Rows
$rows.Add(@('方案', '说明文字怎么生成', '数据是否出内网', '业务假设', '成本构成', '年成本(元)', '三年成本(元)', '优点', '不选的原因'))
$rows.Add(@('一、云端 API（推荐）', '全部报告调 DeepSeek Flash 生成', '脱敏后出内网（待确认）', '月 40 份报告；上下文约 4.4K Token；无并发压力', '纯 Token 费，无固定成本', (R2 $schemeAYear), (R2 ($schemeAYear * 3)), '固定成本为零，零运维，模型质量随厂商升级', '依赖“脱敏后可出内网”这一前提，须信息化负责人书面确认'))
$rows.Add(@('二、本地自建', '内网部署 7B 开源模型（INT8）生成', '不出内网', '峰值并发 2 路；上下文 8K；单卡 RTX 4090 24G', ('自购工作站 ' + $gpuBuy + ' 元按 3 年折旧 + 年电费 + 运维 ' + $modelOps + ' 元/年；部署 ' + $deployDaysB + ' 人天'), (R2 $fixedYearB), (R2 $threeYearB), '数据不出内网，断网可用', '年固定成本约为 API 方案的七千倍，而实际月 Token 量仅为盈亏平衡点的 0.01%'))
$rows.Add(@('三、混合', '固定模板为底，仅“冲突 / 需审批”类报告（约 30%）调 API', '仅该 30% 脱敏后出内网', '同方案一', 'API 费用 × 30%', (R2 $schemeCYear), (R2 ($schemeCYear * 3)), '成本最低，API 故障时仍能出报告', '多数报告的可读性不如模型生成；适合作为方案一未获确认时的过渡'))
$rows.Add(@('盈亏平衡点', ('自建月固定成本 ' + (N0 (R0 $fixedMonthB)) + ' 元 ÷ API 综合单价 ' + (R2 $apiBlend) + ' 元/百万 Token ≈ ' + ($critTok / 1e8).ToString('0.0', $inv) + ' 亿 Token/月'), '', ('本项目实际用量约 ' + (N0 (R0 $tokBaseTokMonth)) + ' Token/月，远低于临界值'), '', '', '', '', '低于临界用量调 API 更便宜；高于才轮到自建，且还要先过合规与运维能力两关'))
Write-Csv '04-部署方案对比表.csv' $rows

$rows = New-Rows
$rows.Add(@('项目', '配置', '测算依据'))
$rows.Add(@('推荐方案', '方案一：云端 API', '数据脱敏后可出内网时采用；未确认前按方案三（固定模板）交付，确认后开通'))
$rows.Add(@('核对程序运行环境', '客户内网虚拟机 4 核 8G（或任意可运行 Python 3.10+ 的机器）×1', '纯标准库 Python，单次核对秒级完成，无并发压力；客户自备'))
$rows.Add(@('模型', 'DeepSeek deepseek-flash（云端 API）', '输入 1.08 / 输出 4.32 元每百万 Token（离峰价），高峰 ×2'))
$rows.Add(@('显存需求', '不适用（不部署本地模型）', '推荐方案无 GPU'))
$rows.Add(@('备选：内网 7B 模型（数据不可出内网时）', '单卡 RTX 4090 24G ×1；系统内存 32G；存储 1TB SSD', ('权重 7 GB（7B × INT8 1 字节）+ KV Cache ' + (R2 $kvGb) + ' GB（56 KB × 8192 × 2 路）= ' + (R2 ($weightGb + $kvGb)) + ' GB，× 1.15 开销 ≈ ' + (R2 $vramGb) + ' GB，单卡 24G 可承载')))
$rows.Add(@('备选：网络与冗余', '单机版；高可用需第二台同配置机器', '月 40 份报告的用量不需要高可用'))
$rows.Add(@('用量包', ('月均 40 次请求、约 ' + (N0 (R0 $tokBaseTokMonth)) + ' Token 内，包干含 20% 缓冲'), '03 表基准档 × 1.2；超量按实测 Token 单价结算'))
Write-Csv '04-推荐配置单.csv' $rows

# ================================================================ 05 ROI 一页纸
$rows = New-Rows
$rows.Add(@('类别', '项目', '计算式', '数值', '单位', '数据来源与假设'))
$rows.Add(@('投入', '一次性投入（实施费）', '06 报价汇总表', $I1, '元', '报价单总价'))
$rows.Add(@('投入', '年运行成本（年服务费）', '06 报价汇总表', $runCost, '元/年', '含运维与模型用量包'))
$rows.Add(@('投入', '首年总投入', '一次性投入 + 年运行成本', $firstTotal, '元', ''))
$rows.Add(@('收益 1 协调工时', '每单协调工时（变更前）', '销售 0.75 + 计划 1.0 + 客服 0.5 + 工厂 0.25 小时', $hoursPerOrder, '小时/单', '假设：访谈有“反复向计划、客服、工厂确认”“同单至少追问两轮”“最花时间是判断版本与核对产能”的陈述，但没有逐环节耗时记录'))
$rows.Add(@('收益 1 协调工时', '工时压降比例', '核对由助手完成后，人工只做复核', ($hoursCut * 100), '%', '假设：试点期实测后更新'))
$rows.Add(@('收益 1 协调工时', '每月节省工时', ($ordersPerMonth.ToString() + ' 单 × ' + $hoursPerOrder + ' 小时 × ' + ($hoursCut * 100).ToString('0', $inv) + '%'), (R2 $hoursSavedMonth), '小时/月', ('约合 ' + $fteSaved.ToString('0.00', $inv) + ' 个全职人力')))
$rows.Add(@('收益 1 协调工时', '节省工时占比', '每月节省工时 ÷（6 人 × 2000 小时/年 ÷ 12）', ($saveRatio * 100).ToString('0.0', $inv), '%', '受影响 6 人为假设'))
$rows.Add(@('收益 1 协调工时', '综合时薪', '月薪 8,000 × 12 × 1.4 ÷ 2000', (R2 $hourCost), '元/小时', '月薪为假设；用综合成本，不用月薪'))
$rows.Add(@('收益 1 协调工时', '协调工时收益（已乘转化系数 0.5）', '每月节省工时 × 12 × 综合时薪 × 0.5', (R0 $gainLabor), '元/年', '等效人力口径，不等于裁人'))
$rows.Add(@('收益 2 产能与时效', 'ERP 录入延迟、缺口显式化', '不折钱', 0, '元/年', '35 分钟延迟与 100 件缺口有访谈依据，但量化不成金额，只写时效改善'))
$rows.Add(@('收益 3 避免冲突', '年版本冲突次数', '现场按旧版做的事件', $errCount, '次/年', '假设：案例主题就是“客户改了订单，现场还在按旧版做”，但没有历史次数'))
$rows.Add(@('收益 3 避免冲突', '单次冲突损失', '返工工时 + 材料报废 + 急单加班 + 客户补偿', $errLoss, '元/次', '假设：按 1,000 件级订单估，签约前须向客户索取返工与补偿记录'))
$rows.Add(@('收益 3 避免冲突', '冲突降低比例', '版本、产能、质量三道核对后', ($errCut * 100), '%', '假设'))
$rows.Add(@('收益 3 避免冲突', '避免损失收益', ('' + $errCount + ' 次/年 × ' + (N0 $errLoss) + ' 元 × ' + ($errCut * 100).ToString('0', $inv) + '%'), (R0 $gainErr), '元/年', '年收益的大头，也是最需要客户数据坐实的一项'))
$rows.Add(@('收益 4 收入增量', '不计', '—', 0, '元/年', '无法验证，不计入'))
$rows.Add(@('结果', '年收益合计（基准）', '人力节省 + 差错减少', (R0 $benefit), '元/年', ''))
$rows.Add(@('结果', '年净收益', '年收益 − 年运行成本', (R0 ($benefit - $runCost)), '元/年', ''))
$rows.Add(@('结果', '首年 ROI', '(年收益 − 首年总投入) ÷ 首年总投入', (Pct $roiBase.roi1), '', ''))
$rows.Add(@('结果', '回本周期（课程公式）', '首年总投入 ÷（年收益 ÷ 12）', (Mon $roiBase.payback), '月', '按年收益（未扣运行成本）摊回'))
$rows.Add(@('结果', '回本点（净现金流口径）', '一次性投入 ÷ 月净收益', (Mon $roiBase.paybackCurve), '月', '回本曲线穿过 0 的月份，比课程公式晚约 2 个月'))
$rows.Add(@('结果', '三年 ROI', '(3 年收益 − 3 年总成本) ÷ 3 年总成本', (Pct $roiBase.three), '', ''))
$rows.Add(@('结果', '续年投入产出比', '年收益 ÷ 年运行成本', ($roiBase.ratio.ToString('0.0', $inv) + ' 倍'), '', '首年之后每年只付年服务费，这是客户每年实际看到的账'))
$rows.Add(@('敏感性', '保守（收益 −30%）：首年 ROI / 回本 / 三年 ROI', '', ((Pct $roiLow.roi1) + ' / ' + (Mon $roiLow.payback) + ' 个月 / ' + (Pct $roiLow.three)), '', '回本为课程公式口径'))
$rows.Add(@('敏感性', '基准：首年 ROI / 回本 / 三年 ROI', '', ((Pct $roiBase.roi1) + ' / ' + (Mon $roiBase.payback) + ' 个月 / ' + (Pct $roiBase.three)), '', ''))
$rows.Add(@('敏感性', '乐观（收益 +30%）：首年 ROI / 回本 / 三年 ROI', '', ((Pct $roiHigh.roi1) + ' / ' + (Mon $roiHigh.payback) + ' 个月 / ' + (Pct $roiHigh.three)), '', ''))
$rows.Add(@('敏感性', '仅工时口径（不计避免冲突）：年净收益', ('' + (R0 $gainLabor) + ' − ' + (N0 $runCost)), (R0 ($gainLabor - $runCost)), '元/年', '年净收益为负、不回本——与阶段 04 的结论一致：只算工时撑不起这个报价'))
foreach ($m in @(0, 6, 12, 18, 24, 30, 36)) {
    $rows.Add(@('回本曲线', ('第 ' + $m + ' 月末累计净现金流'), ('−' + (N0 $I1) + ' + ' + (N0 (R0 $roiBase.netMonth)) + ' × ' + $m), (R0 (-$I1 + $roiBase.netMonth * $m)), '元', '基准情景'))
}
$rows.Add(@('回本前提', '首年回本所需年收益', '首年总投入', $firstTotal, '元/年', '下面四行回答：要么业务量够大，要么冲突损失够大，二者有其一即可首年回本'))
$rows.Add(@('回本前提', '仅靠工时：每月需节省工时', '首年回本所需年收益 ÷ 12 ÷（综合时薪 × 0.5）', (R0 $beHours), '小时/月', ('现假设 ' + (R0 $hoursSavedMonth) + ' 小时/月')))
$rows.Add(@('回本前提', '仅靠工时：对应月单量', '每月需节省工时 ÷（每单协调工时 × 压降比例）', (R0 $beOrders), '单/月', ('现假设 ' + $ordersPerMonth + ' 单/月；客服日均 45 条咨询若多数是改单，真实单量可能远高于此，须抽样核实')))
$rows.Add(@('回本前提', '仅靠避免冲突：对应单次冲突损失', '首年回本所需年收益 ÷（年冲突次数 × 降低比例）', (R0 $beLoss), '元/次', ('现假设 ' + (N0 $errLoss) + ' 元/次')))
$rows.Add(@('回本前提', '仅靠避免冲突：对应年冲突次数', '首年回本所需年收益 ÷（单次冲突损失 × 降低比例）', (R1 $beCount), '次/年', ('现假设 ' + $errCount + ' 次/年')))
Write-Csv '05-客户ROI一页纸.csv' $rows

# ================================================================ 06 报价汇总表
$rows = New-Rows
$rows.Add(@('序号', '项目', '说明', '金额(元)', '计费周期'))
$rows.Add(@(1, '需求调研与方案设计', '业务流程梳理、需求规格、验收指标、ROI 测算报告', '含在实施费', '—'))
$rows.Add(@(2, '系统开发与集成（一次性实施费）', ('' + $totalDays + ' 人天（' + $sumDays + ' 人天工作包 + 20% 缓冲）+ 开发测试环境 + 差旅；成本 × (1 + 15%) × (1 + 25%)'), $implFee, '首年（一次性）'))
$rows.Add(@(3, '年服务费', ('运维 10 人天/年 + 模型用量包（月均 40 次、约 ' + (N0 (R0 $tokBaseTokMonth)) + ' Token 内，含 20% 缓冲），超量按实测 Token 单价结算'), $serviceFee, '首年及续年每年'))
$rows.Add(@(4, '可选：内网 7B 模型部署', ('数据不可出内网时启用：追加 ' + $deployDaysB + ' 人天 + 单卡工作站，详见 04 部署方案对比表'), '另行报价', '—'))
$rows.Add(@('', '首年合计', '2 + 3', $firstYearQuote, '首年'))
$rows.Add(@('', '续年合计', '3', $serviceFee, '续年每年'))
$rows.Add(@('付款节点', '签约 40% / 上线 40% / 验收 20%', '针对一次性实施费；年服务费按年预付', '', ''))
$rows.Add(@('SLA', '工作日 4 小时内响应；结论错误 1 个工作日内修复', '工具为内网离线运行，不承诺在线可用性，承诺每次核对产出审计清单', '', ''))
$rows.Add(@('超出范围另计', '新增系统对接 3 人天/系统起；超出 2 轮的调优按 2,450 元/天；ERP / MES 写接口开发不在范围', '', '', ''))
Write-Csv '06-报价汇总表.csv' $rows

# ================================================================ 咬合自检
$checks = @(
    @{ n = '人天：02 表合计 = 01 表人力小计天数'; ok = ($totalDays -eq $bufDays + $sumDays) }
    @{ n = '人力成本：02 表合计 = 工作包小计 + 缓冲'; ok = ($laborCostTotal -eq $laborCost + $bufCost) }
    @{ n = '人天单价：对外口径 = 成本口径 × 1.25（FDE = 测算器导出的 1,960 / 2,450）'; ok = ($cost['FDE 工程师'] -eq 1960 -and $price['FDE 工程师'] -eq 2450) }
    @{ n = '盈亏平衡：实际月 Token 量 < 临界月 Token 量（推荐 API 成立）'; ok = ($tokBaseTokMonth -lt $critTok) }
    @{ n = '模型费：03 表基准档年成本 × 1.2 + 运维 = 01 表年服务费成本基数'; ok = ([math]::Abs($annualServiceCost - ($opsCost + $tokBaseYear * 1.2)) -lt 0.001) }
    @{ n = '算力：04 表推荐方案无 GPU，01 表服务器块不含 GPU'; ok = ($prodEnv -eq 0) }
    @{ n = '首年合计 = 实施费 + 年服务费'; ok = ($firstYearQuote -eq $implFee + $serviceFee) }
    @{ n = '05 表首年总投入 = 06 表首年合计'; ok = ($firstTotal -eq $firstYearQuote) }
)
Write-Host ''
foreach ($c in $checks) { Write-Host ((@('通过', '未通过')[[int](-not $c.ok)]) + ' | ' + $c.n) }
Write-Host ''
Write-Host ('人天合计 ' + $totalDays + '（工作包 ' + $sumDays + ' + 缓冲 ' + $bufDays + '）；联调占比 ' + ($intShare * 100).ToString('0.0', $inv) + '%')
Write-Host ('人力成本 ' + (N0 $laborCostTotal) + '；加权成本单价 ' + $wc + '；对外口径合计 ' + (N0 $laborPriceTotal))
Write-Host ('单次请求成本 ' + (R2 $perReq) + ' 元；Token 月/年（基准）' + (R2 $tokBaseMonth) + ' / ' + (R2 $tokBaseYear) + ' 元；月 Token ' + (N0 (R0 $tokBaseTokMonth)))
Write-Host ('首年成本 ' + (R2 $firstYearCost) + '；续年成本 ' + (R2 $renewCost))
Write-Host ('实施费 ' + (N0 $implFee) + '；年服务费 ' + (N0 $serviceFee) + '；首年合计 ' + (N0 $firstYearQuote))
Write-Host ('自建：显存 ' + (R2 $vramGb) + ' GB；年固定 ' + (R2 $fixedYearB) + '；三年 ' + (R2 $threeYearB) + '；盈亏平衡 ' + ($critTok / 1e8).ToString('0.0', $inv) + ' 亿 Token/月')
Write-Host ('ROI：每月节省 ' + (R2 $hoursSavedMonth) + ' 小时（' + $fteSaved.ToString('0.00', $inv) + ' 个全职，占 ' + ($saveRatio * 100).ToString('0.0', $inv) + '%）；协调工时 ' + (N0 $gainLabor) + '；避免冲突 ' + (N0 $gainErr) + '；年收益 ' + (N0 $benefit))
Write-Host ('ROI：首年 ' + (Pct $roiBase.roi1) + '；回本 ' + (Mon $roiBase.payback) + ' 个月（曲线 ' + (Mon $roiBase.paybackCurve) + '）；三年 ' + (Pct $roiBase.three) + '；续年投入产出比 ' + $roiBase.ratio.ToString('0.0', $inv) + ' 倍')
Write-Host ('保守 ' + (Pct $roiLow.roi1) + ' / ' + (Mon $roiLow.payback) + ' / ' + (Pct $roiLow.three) + '；乐观 ' + (Pct $roiHigh.roi1) + ' / ' + (Mon $roiHigh.payback) + ' / ' + (Pct $roiHigh.three))
Write-Host ('回本前提：仅靠工时需 ' + (R0 $beOrders) + ' 单/月（现 ' + $ordersPerMonth + '）；仅靠避免冲突需 ' + (R0 $beLoss) + ' 元/次（现 ' + $errLoss + '）或 ' + (R1 $beCount) + ' 次/年（现 ' + $errCount + '）')

# ---------------------------------------------------------------- Excel 工作簿 + Markdown 版
$wbScript = Join-Path $PSScriptRoot 'build_workbook.ps1'
if (Get-Module -ListAvailable ImportExcel) { . $wbScript }
else { Write-Host '未安装 ImportExcel 模块，跳过 Excel 工作簿与 Markdown 版的生成' }
