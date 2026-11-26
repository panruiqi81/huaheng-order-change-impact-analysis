# 由 build_pricing.ps1 末尾调用（共用其变量与 $Tables），生成：
#   1) 华衡智造-企业报价包.xlsx —— 每件一个工作表，汇总格为引用公式，参数集中在“参数”表
#   2) 报价包.md —— GitHub 可直接渲染的表格版
$ErrorActionPreference = 'Stop'
Import-Module ImportExcel -ErrorAction Stop

$C_HEAD = [System.Drawing.Color]::FromArgb(0x1F, 0x3B, 0x63)
$C_SUB = [System.Drawing.Color]::FromArgb(0xEA, 0xEF, 0xF6)
$C_QUOTE = [System.Drawing.Color]::FromArgb(0xFF, 0xF2, 0xCC)
$C_LINE = [System.Drawing.Color]::FromArgb(0xD0, 0xD7, 0xE2)
$C_GREY = [System.Drawing.Color]::FromArgb(0x59, 0x59, 0x59)
$C_ACC = [System.Drawing.Color]::FromArgb(0xB0, 0x30, 0x20)
$FONT = '微软雅黑'

function Cell { param([object]$v, [string]$fmt = '') return @{ v = $v; f = $null; fmt = $fmt } }
function Fx { param([string]$f, [string]$fmt = '') return @{ v = $null; f = $f; fmt = $fmt } }
function Row { param([string]$k, [object[]]$c, [int[]]$m = $null) return @{ k = $k; c = $c; m = $m } }
function IsNum([object]$v) { return ($null -ne $v) -and ($v -is [byte] -or $v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) }
function AutoCell([object]$v) {
    if ($null -eq $v) { return (Cell '') }
    if (IsNum $v) {
        if ([math]::Abs([double]$v - [math]::Round([double]$v)) -gt 1e-9) { return (Cell $v '#,##0.00') }
        return (Cell $v '#,##0')
    }
    if ($v -is [string] -and $v -match '^-?\d+(\.\d+)?$') {
        $d = [double]$v
        if ($v.Contains('.')) { return (Cell $d '#,##0.00') }
        return (Cell $d '#,##0')
    }
    return (Cell $v)
}
function VisLen([string]$s) {   # 显示宽度：中日韩字符按 2 计
    $n = 0
    foreach ($ch in $s.ToCharArray()) { if ([int]$ch -gt 0x2E7F) { $n += 2 } else { $n += 1 } }
    return $n
}
function Convert-Table([object[]]$t, [hashtable]$kinds) {
    $out = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $t.Count; $i++) {
        $kind = 'data'; if ($i -eq 0) { $kind = 'header' }; if ($kinds.ContainsKey($i)) { $kind = $kinds[$i] }
        $cells = [System.Collections.Generic.List[object]]::new()
        foreach ($v in $t[$i]) { $cells.Add((AutoCell $v)) }
        $out.Add((Row $kind $cells.ToArray()))
    }
    return ,$out
}
function With-Header([string]$title, [string]$note, [object[]]$body) {
    $out = [System.Collections.Generic.List[object]]::new()
    $out.Add((Row 'title' @((Cell $title))))
    if ($note) { $out.Add((Row 'note' @((Cell $note)))) }
    foreach ($x in $body) { $out.Add($x) }
    return ,$out
}
function Pad-Row([string]$k, [object[]]$cells, [int]$n) {
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($c in $cells) { $list.Add($c) }
    while ($list.Count -lt $n) { $list.Add((Cell '')) }
    return (Row $k $list.ToArray())
}
function Add-Sheet($Excel, [string]$Name, [object[]]$Rows, [double[]]$Widths, [int]$FreezeAfter = 0) {
    $ws = $Excel.Workbook.Worksheets.Add($Name)
    $ws.View.ShowGridLines = $false
    $ncol = $Widths.Count
    $totalW = 0.0; foreach ($w in $Widths) { $totalW += $w }
    $r = 0
    foreach ($row in $Rows) {
        $r++
        if ($row.k -eq 'blank') { continue }
        $range = $ws.Cells[$r, 1, $r, $ncol]
        $range.Style.Font.Name = $FONT
        $range.Style.Font.Size = 10
        if ($row.k -eq 'title') {
            $ws.Cells[$r, 1].Value = [string]$row.c[0].v
            if ($ncol -gt 1) { $ws.Cells[$r, 1, $r, $ncol].Merge = $true }
            $range.Style.Font.Bold = $true
            $range.Style.Font.Size = 15
            $range.Style.Font.Color.SetColor($C_HEAD)
            $ws.Row($r).Height = 32
        }
        elseif ($row.k -eq 'note') {
            $ws.Cells[$r, 1].Value = [string]$row.c[0].v
            if ($ncol -gt 1) { $ws.Cells[$r, 1, $r, $ncol].Merge = $true }
            $range.Style.Font.Italic = $true
            $range.Style.Font.Size = 9
            $range.Style.Font.Color.SetColor($C_GREY)
            $range.Style.WrapText = $true
            $range.Style.VerticalAlignment = 'Top'
            $lines = [math]::Ceiling(((VisLen ([string]$row.c[0].v)) + $ncol * 2) / [math]::Max(8, $totalW - 6)) + 1
            $ws.Row($r).Height = [double]($lines * 14.5 + 8)
        }
        else {
            $c = 0; $lines = 1
            foreach ($cell in $row.c) {
                $c++
                if ($c -gt $ncol) { break }
                $o = $ws.Cells[$r, $c]
                if ($cell.f) { $o.Formula = $cell.f }
                elseif ($null -ne $cell.v -and [string]$cell.v -ne '') { $o.Value = $cell.v }
                if ($cell.fmt) {
                    $o.Style.Numberformat.Format = $cell.fmt
                    $o.Style.HorizontalAlignment = 'Right'
                    $o.Style.VerticalAlignment = 'Top'
                }
                else {
                    $o.Style.WrapText = $true
                    $o.Style.VerticalAlignment = 'Top'
                }
                if ($null -ne $cell.v -and [string]$cell.v -ne '') {
                    $n = [math]::Ceiling(((VisLen ([string]$cell.v)) + 2) / [math]::Max(4, $Widths[$c - 1] - 1))
                    if ($n -gt $lines) { $lines = $n }
                }
            }
            if ($row.m) {   # 合并区的文字按合并后的总宽度估算行数
                $mw = 0.0; for ($k = $row.m[0]; $k -le $row.m[1]; $k++) { $mw += $Widths[$k - 1] }
                $first = $row.c[$row.m[0] - 1]
                if ($first.fmt -eq '' -and $null -ne $first.v) {
                    $n = [math]::Ceiling(((VisLen ([string]$first.v)) + 2) / [math]::Max(4, $mw - 1))
                    if ($n -gt $lines) { $lines = $n }
                }
            }
            foreach ($side in 'Top', 'Bottom', 'Left', 'Right') {
                $range.Style.Border.$side.Style = 'Thin'
                $range.Style.Border.$side.Color.SetColor($C_LINE)
            }
            switch ($row.k) {
                'header' {
                    $range.Style.Font.Bold = $true
                    $range.Style.Font.Color.SetColor([System.Drawing.Color]::White)
                    $range.Style.Fill.PatternType = 'Solid'
                    $range.Style.Fill.BackgroundColor.SetColor($C_HEAD)
                    $range.Style.HorizontalAlignment = 'Center'
                    $range.Style.VerticalAlignment = 'Center'
                    $range.Style.WrapText = $true
                    $ws.Row($r).Height = 32
                }
                'sub' {
                    $range.Style.Font.Bold = $true
                    $range.Style.Fill.PatternType = 'Solid'
                    $range.Style.Fill.BackgroundColor.SetColor($C_SUB)
                }
                'total' {
                    $range.Style.Font.Bold = $true
                    $range.Style.Fill.PatternType = 'Solid'
                    $range.Style.Fill.BackgroundColor.SetColor($C_SUB)
                    $range.Style.Border.Top.Style = 'Medium'
                }
                'quote' {
                    $range.Style.Font.Bold = $true
                    $range.Style.Fill.PatternType = 'Solid'
                    $range.Style.Fill.BackgroundColor.SetColor($C_QUOTE)
                }
                'section' {
                    $ws.Cells[$r, 1].Style.Font.Bold = $true
                    $ws.Cells[$r, 1].Style.Font.Color.SetColor($C_ACC)
                }
                default { }
            }
            if ($row.m) { $ws.Cells[$r, $row.m[0], $r, $row.m[1]].Merge = $true }
            if ($lines -gt 1) { $ws.Row($r).Height = [double]($lines * 14.5 + 6) }
        }
    }
    for ($i = 1; $i -le $ncol; $i++) { $ws.Column($i).Width = $Widths[$i - 1] }
    if ($FreezeAfter -gt 0) { $ws.View.FreezePanes($FreezeAfter + 1, 1) }
    try {   # 打印：横向、适应一页宽
        $ws.PrinterSettings.Orientation = 'Landscape'
        $ws.PrinterSettings.FitToPage = $true
        $ws.PrinterSettings.FitToWidth = 1
        $ws.PrinterSettings.FitToHeight = 0
        $ws.PrinterSettings.RepeatRows = '1:3'
    } catch { }
    return $ws
}

# ---------------------------------------------------------------- 参数表（先定行号，供各表公式引用）
function PDef { param($k, $n, $v, $fmt, $s) return @{ k = $k; n = $n; v = $v; fmt = $fmt; s = $s } }
$paramDefs = @(
    (PDef 'coef' '综合成本系数' $coef '0.00' '社保公积金、场地设备、管理分摊的放大系数')
    (PDef 'util' '可计费利用率' $util '0%' '工程师可计费时间占比')
    (PDef 'dpm' '每月可计费工作日' $dpm '0.00' '')
    (PDef 'markup' '成本加成率' $markup '0%' '对外人天单价 = 人天综合成本 × (1 + 加成率)')
    (PDef 'risk' '风险预留率' $risk '0%' '需求越模糊越高')
    (PDef 'buf' '整体缓冲率' $bufRate '0%' '加在汇总工时上，单独成行')
    (PDef 'reports' '月核对报告量（份/月）' $ordersPerMonth '0' '假设，试点期实测后更新')
    (PDef 'steps' 'Agent 步数' $steps '0' '每次请求调用模型几步')
    (PDef 'retry' '重试率' $retry '0%' '格式校验不通过重生成')
    (PDef 'cache' '输入缓存命中率' $cacheHit '0%' '按 0% 保守计')
    (PDef 'pIn' '输入单价（元/百万 Token）' $pIn '0.00' 'DeepSeek Flash 离峰价')
    (PDef 'pOut' '输出单价（元/百万 Token）' $pOut '0.00' 'DeepSeek Flash 离峰价')
    (PDef 'outTok' '单次输出 Token' $outTok '#,##0' '')
    (PDef 'packBuf' '模型用量缓冲率' 0.20 '0%' '包干量含 20% 缓冲')
    (PDef 'devVm' '开发测试云主机（元/月）' 120 '#,##0' '假设，待按云平台实时报价复核')
    (PDef 'devMonths' '开发测试环境使用月数' 3 '0' '')
    (PDef 'travelN' '差旅次数' 4 '0' '')
    (PDef 'travelOnce' '单次差旅（元）' 1500 '#,##0' '假设')
    (PDef 'opsRespM' '日常响应与巡检（人天/月）' 0.5 '0.0' '')
    (PDef 'opsIter' '版本迭代（人天/年）' 4 '0' '')
    (PDef 'gpu' '内网工作站采购价（元）' 25000 '#,##0' '单卡 RTX 4090 24G，假设')
    (PDef 'depYears' '折旧年限（年）' 3 '0' '')
    (PDef 'kw' '整机功耗（kW）' 0.6 '0.0' '')
    (PDef 'elec' '电价（元/度）' 0.8 '0.00' '')
    (PDef 'pue' 'PUE' 1.4 '0.0' '机房能效系数')
    (PDef 'modelOps' '模型服务运维（元/月）' 1000 '#,##0' '')
    (PDef 'deployDays' '内网模型部署（人天）' 5 '0' '')
    (PDef 'mix' '混合方案调 API 的报告比例' $mixShare '0%' '')
    (PDef 'weightGb' '模型权重显存（GB）' $weightGb '0' '7B × INT8 1 字节')
    (PDef 'kvKb' '每 Token KV 占用（KB）' $kvKb '0' '7B 模型典型值')
    (PDef 'ctx' '上下文长度（Token）' $ctx '#,##0' '')
    (PDef 'conc' '峰值并发（路）' $conc '0' '')
    (PDef 'overhead' '显存运行开销' 1.15 '0.00' '')
    (PDef 'people' '受影响岗位人数' $people '0' '假设')
    (PDef 'salary' '岗位月薪（元）' $monthSalary '#,##0' '假设')
    (PDef 'hoursOrder' '每单协调工时（变更前，小时）' $hoursPerOrder '0.00' '销售 / 计划 / 客服 / 工厂合计；假设，试点期抽样计时后更新')
    (PDef 'cut' '协调工时压降比例' $hoursCut '0%' '核对由助手完成后人工只做复核；假设')
    (PDef 'conv' '人力节省转化系数' $conv '0%' '省下的工时不等于省下的钱')
    (PDef 'errN' '年版本冲突次数' $errCount '0' '假设，案例无历史数据')
    (PDef 'errLoss' '单次冲突损失（元）' $errLoss '#,##0' '假设，案例无历史数据')
    (PDef 'errCut' '冲突降低比例' $errCut '0%' '假设')
    (PDef 'sens' '敏感性幅度' 0.3 '0%' '收益 ±30%')
)
$P = @{}
for ($i = 0; $i -lt $paramDefs.Count; $i++) { $P[$paramDefs[$i].k] = 4 + $i }   # 标题 1、说明 2、表头 3
function PN([string]$k) { return ('参数!$B${0}' -f $P[$k]) }

# ---------------------------------------------------------------- 跨表引用前缀与关键单元格
$S1 = "'1-成本拆解表'!"; $S2 = "'2-人天估算与单价表'!"; $S2b = "'2-人天单价推导'!"; $S3 = "'3-Token月成本测算表'!"; $S5 = "'5-客户ROI一页纸'!"; $S6 = "'6-报价汇总表'!"
# Token 表：标题 1、说明 2、构成表头 3、构成 4-9、输入合计 10、输出 11、三档表头 13、档位 14-16
$T_IN = $S3 + 'B10'; $T_OUT = $S3 + 'B11'; $T_YEAR = $S3 + 'O15'; $T_TOK = $S3 + 'P15'
$T_PRICE_IN = $S3 + 'K14'; $T_PRICE_OUT = $S3 + 'L14'

# ---------------------------------------------------------------- 新建工作簿
$xlsx = Join-Path $OutDir '华衡智造-企业报价包.xlsx'
if (Test-Path -LiteralPath $xlsx) { Remove-Item -LiteralPath $xlsx }
$xl = Export-Excel -Path $xlsx -WorksheetName 'tmp' -InputObject ([pscustomobject]@{ a = 1 }) -PassThru
$xl.Workbook.Worksheets.Delete('tmp')
$roleKeys = @($ref.Keys)
$fdeRow = 4 + [array]::IndexOf($roleKeys, 'FDE 工程师')
$lastTask = 3 + $tasks.Count; $rSub = $lastTask + 1; $rBuf = $rSub + 1; $rTot = $rBuf + 1   # 人天表：表头 3、工作包 4..21、小计 22、缓冲 23、合计 24

# ---- 2-人天单价推导（角色 3-7 行）
$t2b = Convert-Table $Tables['02-人天单价推导.csv'] @{ 6 = 'data'; 7 = 'note' }
for ($i = 0; $i -lt $roleKeys.Count; $i++) {
    $er = 4 + $i; $row = $t2b[1 + $i]
    $row.c[2] = Fx ('={0}' -f (PN 'coef')) '0.00'
    $row.c[3] = Fx ('={0}' -f (PN 'util')) '0%'
    $row.c[4] = Fx ('={0}' -f (PN 'markup')) '0%'
    $row.c[5] = Fx ('=B{0}*C{0}' -f $er) '#,##0'
    $row.c[6] = Fx ('=ROUND(F{0}/{1}/D{0},0)' -f $er, (PN 'dpm')) '#,##0'
    $row.c[7] = Fx ('=ROUND(G{0}*(1+E{0}),0)' -f $er) '#,##0'
}
$rows2b = With-Header '人天单价推导' '人天综合成本 = 月薪 × 综合成本系数 ÷ 21.75 ÷ 利用率；对外人天单价 = 人天综合成本 × (1 + 成本加成率)。FDE 一行与测算器导出结果一致（1,960 / 2,450）。改“参数”表即全表联动。' $t2b
Add-Sheet $xl '2-人天单价推导' $rows2b @(22, 13, 13, 14, 13, 14, 15, 15, 16) 3 | Out-Null

# ---- 2-人天估算与单价表（工作包 3-20 行）
$t2a = Convert-Table $Tables['02-人天估算与单价表.csv'] @{ 19 = 'sub'; 20 = 'sub'; 21 = 'total' }
for ($i = 0; $i -lt $tasks.Count; $i++) {
    $er = 4 + $i; $row = $t2a[1 + $i]; $rr = 4 + [array]::IndexOf($roleKeys, $tasks[$i].r)
    $row.c[5] = Fx ('={0}G{1}' -f $S2b, $rr) '#,##0'
    $row.c[6] = Fx ('={0}H{1}' -f $S2b, $rr) '#,##0'
    $row.c[7] = Fx ('=E{0}*F{0}' -f $er) '#,##0'
    $row.c[8] = Fx ('=E{0}*G{0}' -f $er) '#,##0'
}
$row = $t2a[1 + $tasks.Count]      # 小计
$row.c[4] = Fx ('=SUM(E4:E{0})' -f $lastTask) '0'
$row.c[5] = Fx ('=ROUND(H{0}/E{0},0)' -f $rSub) '#,##0'
$row.c[6] = Fx ('=ROUND(I{0}/E{0},0)' -f $rSub) '#,##0'
$row.c[7] = Fx ('=SUM(H4:H{0})' -f $lastTask) '#,##0'
$row.c[8] = Fx ('=SUM(I4:I{0})' -f $lastTask) '#,##0'
$row = $t2a[2 + $tasks.Count]      # 缓冲
$row.c[4] = Fx ('=ROUND(E{0}*{1},0)' -f $rSub, (PN 'buf')) '0'
$row.c[5] = Fx ('=F{0}' -f $rSub) '#,##0'
$row.c[6] = Fx ('=G{0}' -f $rSub) '#,##0'
$row.c[7] = Fx ('=E{0}*F{0}' -f $rBuf) '#,##0'
$row.c[8] = Fx ('=E{0}*G{0}' -f $rBuf) '#,##0'
$row = $t2a[3 + $tasks.Count]      # 合计
$row.c[4] = Fx ('=E{0}+E{1}' -f $rSub, $rBuf) '0'
$row.c[7] = Fx ('=H{0}+H{1}' -f $rSub, $rBuf) '#,##0'
$row.c[8] = Fx ('=I{0}+I{1}' -f $rSub, $rBuf) '#,##0'
$rows2a = With-Header '人天估算与单价表' '18 个工作包，角色 × 天数 × 单价。单价取自“2-人天单价推导”表；报价以“成本口径”一列为基数（见“1-成本拆解表”说明）。' $t2a
Add-Sheet $xl '2-人天估算与单价表' $rows2a @(26, 6, 52, 17, 7, 12, 12, 14, 14, 38) 3 | Out-Null

# ---- 3-Token月成本测算表
$tokW = @(30, 22, 40, 11, 12, 14, 14, 10, 10, 14, 16, 16, 13, 12, 12, 14)
$tokX = [System.Collections.Generic.List[object]]::new()
$tokX.Add((Row 'title' @((Cell 'Token 月成本测算表'))))
$tokX.Add((Row 'note' @((Cell '模型 DeepSeek deepseek-flash，离峰价（高峰时段 ×2）。单次请求成本 =（输入 × 输入单价 + 输出 × 输出单价）× Agent 步数 ×（1 + 重试率）；月成本 = 单次请求成本 × 日均请求数 × 30；年成本 = 月成本 × 12。单次 Token 为估算值，上线前用真实报告样本实测后重算。'))))
$th = Pad-Row 'header' @((Cell '单次输入构成'), (Cell 'Token'), (Cell '说明')) 16
$th.m = @(4, 16)
$tokX.Add($th)
for ($i = 0; $i -lt $inParts.Count; $i++) {
    $pr = Pad-Row 'data' @((Cell $inParts[$i].n), (Cell $inParts[$i].t '0'), (Cell $inParts[$i].s)) 16
    $pr.m = @(3, 16)
    $tokX.Add($pr)
}
$tr = Pad-Row 'sub' @((Cell '单次输入合计'), (Fx '=SUM(B4:B9)' '#,##0'), (Cell '系统提示词 + 结构化结果 + 规则依据 + 订单号与指令')) 16
$tr.m = @(3, 16); $tokX.Add($tr)
$tr = Pad-Row 'sub' @((Cell '单次输出'), (Fx ('={0}' -f (PN 'outTok')) '#,##0'), (Cell '约 1,000 字中文说明')) 16
$tr.m = @(3, 16); $tokX.Add($tr)
$tokX.Add((Row 'blank' @()))
$tokX.Add((Row 'header' @((Cell '档位'), (Cell '模型'), (Cell '业务量系数'), (Cell '日均请求数(次)'), (Cell '月请求数(次)'), (Cell '单次输入Token'), (Cell '单次输出Token'), (Cell 'Agent步数'), (Cell '重试率(%)'), (Cell '输入缓存命中率(%)'), (Cell '输入单价(元/百万Token)'), (Cell '输出单价(元/百万Token)'), (Cell '单次请求成本(元)'), (Cell '月成本(元)'), (Cell '年成本(元)'), (Cell '月Token总量'))))
for ($i = 0; $i -lt $levels.Count; $i++) {
    $er = 14 + $i
    $c = [System.Collections.Generic.List[object]]::new()
    $c.Add((Cell $levels[$i].name)); $c.Add((Cell 'DeepSeek deepseek-flash')); $c.Add((Cell $levels[$i].f '0.0'))
    $c.Add((Fx ('={0}/30' -f (PN 'reports')) '0.00'))
    $c.Add((Fx ('=D{0}*30*C{0}' -f $er) '#,##0.0'))
    $c.Add((Fx ('={0}' -f $T_IN) '#,##0')); $c.Add((Fx ('={0}' -f $T_OUT) '#,##0'))
    $c.Add((Fx ('={0}' -f (PN 'steps')) '0')); $c.Add((Fx ('={0}' -f (PN 'retry')) '0%')); $c.Add((Fx ('={0}' -f (PN 'cache')) '0%'))
    $c.Add((Fx ('={0}' -f (PN 'pIn')) '0.00')); $c.Add((Fx ('={0}' -f (PN 'pOut')) '0.00'))
    $c.Add((Fx ('=(F{0}*K{0}+G{0}*L{0})/1000000*H{0}*(1+I{0})' -f $er) '0.0000'))
    $c.Add((Fx ('=M{0}*E{0}' -f $er) '0.00')); $c.Add((Fx ('=N{0}*12' -f $er) '0.00')); $c.Add((Fx ('=E{0}*(F{0}+G{0})' -f $er) '#,##0'))
    $kind = 'data'; if ($levels[$i].name -eq '基准') { $kind = 'sub' }
    $tokX.Add((Row $kind $c.ToArray()))
}
$tokX.Add((Row 'note' @((Cell '基准档年成本即“6-报价汇总表”年服务费里的模型费；报价包干量含 20% 用量缓冲，超量按实测 Token 单价结算。'))))
Add-Sheet $xl '3-Token月成本测算表' $tokX $tokW 2 | Out-Null

# ---- 1-成本拆解表（标题 1、说明 2、表头 3；CSV 第 i 行 = 工作表第 i+3 行，数据自第 4 行起）
$t1 = Convert-Table $Tables['01-成本拆解表.csv'] @{ 19 = 'sub'; 20 = 'sub'; 30 = 'total'; 31 = 'total'; 32 = 'quote'; 33 = 'quote'; 34 = 'quote'; 35 = 'section'; 36 = 'section' }
for ($i = 0; $i -lt $tasks.Count; $i++) {
    $er = 4 + $i; $row = $t1[1 + $i]; $rr = 4 + [array]::IndexOf($roleKeys, $tasks[$i].r)
    $row.c[3] = Fx ('={0}E{1}' -f $S2, $er) '0'
    $row.c[4] = Fx ('={0}G{1}' -f $S2b, $rr) '#,##0'
    $row.c[6] = Fx ('=D{0}*E{0}' -f $er) '#,##0'
}
$row = $t1[19]   # 整体缓冲（第 22 行）
$row.c[3] = Fx ('={0}E{1}' -f $S2, $rBuf) '0'; $row.c[4] = Fx ('={0}F{1}' -f $S2, $rBuf) '#,##0'; $row.c[6] = Fx '=D22*E22' '#,##0'
$row = $t1[20]   # 人力小计（第 23 行）
$row.c[3] = Fx '=SUM(D4:D22)' '0'; $row.c[6] = Fx '=SUM(G4:G22)' '#,##0'
$row = $t1[21]; $row.c[6] = Fx ('={0}' -f $T_YEAR) '0.00'                                     # 模型（第 24 行）
$row = $t1[22]; $row.c[6] = Fx ('={0}*{1}' -f $T_YEAR, (PN 'packBuf')) '0.00'
$row = $t1[23]; $row.c[6] = Fx ('={0}*{1}' -f (PN 'devVm'), (PN 'devMonths')) '#,##0'        # 开发测试环境（第 26 行）
$row = $t1[27]   # 日常响应（第 30 行）
$row.c[3] = Fx ('={0}*12' -f (PN 'opsRespM')) '0'; $row.c[4] = Fx ('={0}G{1}' -f $S2b, $fdeRow) '#,##0'; $row.c[6] = Fx '=D30*E30' '#,##0'
$row = $t1[28]   # 版本迭代（第 31 行）
$row.c[3] = Fx ('={0}' -f (PN 'opsIter')) '0'; $row.c[4] = Fx ('={0}G{1}' -f $S2b, $fdeRow) '#,##0'; $row.c[6] = Fx '=D31*E31' '#,##0'
$row = $t1[29]; $row.c[6] = Fx ('={0}*{1}' -f (PN 'travelN'), (PN 'travelOnce')) '#,##0'    # 差旅（第 32 行）
$row = $t1[30]; $row.c[6] = Fx '=G23+G26+G32+SUM(G24:G25)+SUM(G30:G31)' '#,##0.00'          # 首年成本合计（第 33 行）
$row = $t1[31]; $row.c[6] = Fx '=SUM(G24:G25)+SUM(G30:G31)' '#,##0.00'                       # 续年成本合计（第 34 行）
$row = $t1[32]   # 一次性实施费（第 35 行）
$row.c[5] = Fx ('="成本 "&TEXT(G23+G26+G32,"#,##0")&" × (1 + "&TEXT({0},"0%")&") × (1 + "&TEXT({1},"0%")&")"' -f (PN 'risk'), (PN 'markup'))
$row.c[6] = Fx ('=(G23+G26+G32)*(1+{0})*(1+{1})' -f (PN 'risk'), (PN 'markup')) '#,##0'
$row = $t1[33]   # 年服务费（第 36 行）
$row.c[5] = Fx ('="成本 "&TEXT(SUM(G24:G25)+SUM(G30:G31),"#,##0.00")&" × (1 + "&TEXT({0},"0%")&") × (1 + "&TEXT({1},"0%")&")"' -f (PN 'risk'), (PN 'markup'))
$row.c[6] = Fx ('=(SUM(G24:G25)+SUM(G30:G31))*(1+{0})*(1+{1})' -f (PN 'risk'), (PN 'markup')) '#,##0'
$row = $t1[34]   # 首年合计（第 37 行）
$row.c[6] = Fx '=G35+G36' '#,##0'; $row.c[8] = Fx '="续年 = 年服务费 "&TEXT(G36,"#,##0")'
$rows1 = With-Header '成本拆解表' '五块成本逐项列示，金额为成本口径。对外报价 = 成本 × (1 + 风险预留率) × (1 + 成本加成率)。练 2 的对外人天单价已含 25% 加成，故这里以人天综合成本为基数，只加成一次。改“参数”表即全表联动。' $t1
Add-Sheet $xl '1-成本拆解表' $rows1 @(11, 54, 15, 7, 10, 46, 14, 14, 40) 3 | Out-Null

# ---- 6-报价汇总表（表头 3；实施费第 5 行、年服务费第 6 行、首年合计第 8 行、续年第 9 行）
$t6 = Convert-Table $Tables['06-报价汇总表.csv'] @{ 5 = 'total'; 6 = 'total'; 7 = 'section'; 8 = 'section'; 9 = 'section' }
$row = $t6[2]
$row.c[2] = Fx ('={0}E{1}&" 人天（工作包 "&{0}E{2}&" 人天 + "&TEXT({3},"0%")&" 缓冲）+ 开发测试环境 + 差旅；成本 × (1 + 风险预留率) × (1 + 成本加成率)"' -f $S2, $rTot, $rSub, (PN 'buf'))
$row.c[3] = Fx ('={0}G35' -f $S1) '#,##0'
$row = $t6[3]
$row.c[2] = Fx ('="运维 "&TEXT({0}D30+{0}D31,"0")&" 人天/年 + 模型用量包（月均 "&TEXT({1},"0")&" 次、约 "&TEXT({2},"#,##0")&" Token 内，含 "&TEXT({3},"0%")&" 缓冲），超量按实测 Token 单价结算"' -f $S1, (PN 'reports'), $T_TOK, (PN 'packBuf'))
$row.c[3] = Fx ('={0}G36' -f $S1) '#,##0'
$row = $t6[5]; $row.c[3] = Fx '=D5+D6' '#,##0'
$row = $t6[6]; $row.c[3] = Fx '=D6' '#,##0'
$rows6 = With-Header '报价汇总表（正式报价单）' '首年与续年分开；付款节点、SLA、超出范围另计均已写明。金额由“1-成本拆解表”与“3-Token月成本测算表”联动。' $t6
Add-Sheet $xl '6-报价汇总表' $rows6 @(9, 30, 70, 16, 16) 3 | Out-Null

# ---- 5-客户ROI一页纸（表头 3；数据自第 4 行起）
$roiX = [System.Collections.Generic.List[object]]::new()
$roiX.Add((Row 'title' @((Cell '客户 ROI 一页纸'))))
$roiX.Add((Row 'note' @((Cell '站在客户角度算账。首年含一次性实施费，账面为负是正常的；客户真正看的是多久回本、三年赚回多少、以及首年之后每年付 1 元服务费换回几元。'))))
$roiX.Add((Row 'header' @((Cell '类别'), (Cell '项目'), (Cell '计算式'), (Cell '数值'), (Cell '单位'), (Cell '数据来源与假设'))))
function RoiRow([string]$a, [string]$b, [object]$calc, [object]$val, [string]$unit, [object]$src, [string]$kind = 'data') {
    if ($src -isnot [string]) { return (Row $kind @((Cell $a), (Cell $b), $calc, $val, (Cell $unit), $src)) }
    return (Row $kind @((Cell $a), (Cell $b), $calc, $val, (Cell $unit), (Cell $src)))
}
$roiX.Add((RoiRow '投入' '一次性投入（实施费）' (Cell '报价汇总表') (Fx ('={0}D5' -f $S6) '#,##0') '元' '报价单总价'))
$roiX.Add((RoiRow '投入' '年运行成本（年服务费）' (Cell '报价汇总表') (Fx ('={0}D6' -f $S6) '#,##0') '元/年' '含运维与模型用量包'))
$roiX.Add((RoiRow '投入' '首年总投入' (Cell '一次性投入 + 年运行成本') (Fx '=D4+D5' '#,##0') '元' '' 'sub'))
$roiX.Add((RoiRow '收益 1 协调工时' '每单协调工时（变更前）' (Cell '销售 0.75 + 计划 1.0 + 客服 0.5 + 工厂 0.25 小时') (Fx ('={0}' -f (PN 'hoursOrder')) '0.00') '小时/单' '假设：访谈有“反复向计划、客服、工厂确认”“同单至少追问两轮”“最花时间是判断版本与核对产能”，但无逐环节耗时'))
$roiX.Add((RoiRow '收益 1 协调工时' '工时压降比例' (Cell '核对由助手完成后，人工只做复核') (Fx ('={0}' -f (PN 'cut')) '0%') '%' '假设'))
$roiX.Add((RoiRow '收益 1 协调工时' '每月节省工时' (Cell '月单量 × 每单协调工时 × 压降比例') (Fx ('={0}*D7*D8' -f (PN 'reports')) '0.0') '小时/月' (Fx '="约合 "&TEXT(D9/(2000/12),"0.00")&" 个全职人力"')))
$roiX.Add((RoiRow '收益 1 协调工时' '节省工时占比' (Cell '每月节省工时 ÷（人数 × 2000 小时/年 ÷ 12）') (Fx ('=D9/({0}*2000/12)' -f (PN 'people')) '0.0%') '%' '受影响人数为假设'))
$roiX.Add((RoiRow '收益 1 协调工时' '综合时薪' (Cell '月薪 × 12 × 1.4 ÷ 2000') (Fx ('={0}*12*1.4/2000' -f (PN 'salary')) '0.00') '元/小时' '月薪为假设；用综合成本'))
$roiX.Add((RoiRow '收益 1 协调工时' '协调工时收益（已乘转化系数）' (Cell '每月节省工时 × 12 × 综合时薪 × 转化系数') (Fx ('=D9*12*D11*{0}' -f (PN 'conv')) '#,##0') '元/年' '等效人力口径，不等于裁人'))
$roiX.Add((RoiRow '收益 2 产能与时效' '答复提速、缺口显式化' (Cell '不折钱') (Cell 0 '#,##0') '元/年' '35 分钟录入延迟与 100 件缺口有访谈依据，但量化不成金额'))
$roiX.Add((RoiRow '收益 3 避免冲突' '年版本冲突次数' (Cell '现场按旧版做的事件') (Fx ('={0}' -f (PN 'errN')) '0') '次/年' '假设：案例主题就是“客户改了订单，现场还在按旧版做”，但无历史次数'))
$roiX.Add((RoiRow '收益 3 避免冲突' '单次冲突损失' (Cell '返工工时 + 材料报废 + 急单加班 + 客户补偿') (Fx ('={0}' -f (PN 'errLoss')) '#,##0') '元/次' '假设：按 1,000 件级订单估'))
$roiX.Add((RoiRow '收益 3 避免冲突' '冲突降低比例' (Cell '版本、产能、质量三道核对后') (Fx ('={0}' -f (PN 'errCut')) '0%') '%' '假设'))
$roiX.Add((RoiRow '收益 3 避免冲突' '避免损失收益' (Cell '年冲突次数 × 单次损失 × 降低比例') (Fx '=D14*D15*D16' '#,##0') '元/年' '年收益的大头，也是最需要客户数据坐实的一项'))
$roiX.Add((RoiRow '收益 4 收入增量' '不计' (Cell '—') (Cell 0 '#,##0') '元/年' '无法验证，不计入'))
$roiX.Add((RoiRow '结果' '年收益合计（基准）' (Cell '协调工时 + 避免冲突') (Fx '=D12+D13+D17+D18' '#,##0') '元/年' '' 'sub'))
$roiX.Add((RoiRow '结果' '年净收益' (Cell '年收益 − 年运行成本') (Fx '=D19-D5' '#,##0') '元/年' ''))
$roiX.Add((RoiRow '结果' '首年 ROI' (Cell '(年收益 − 首年总投入) ÷ 首年总投入') (Fx '=(D19-D6)/D6' '0.0%') '' '含一次性实施费' 'total'))
$roiX.Add((RoiRow '结果' '回本周期（课程公式）' (Cell '首年总投入 ÷（年收益 ÷ 12）') (Fx '=D6/(D19/12)' '0.0') '月' '按年收益（未扣运行成本）摊回' 'total'))
$roiX.Add((RoiRow '结果' '回本点（净现金流口径）' (Cell '一次性投入 ÷ 月净收益') (Fx '=IF(D20>0,D4/(D20/12),"不回本")' '0.0') '月' '回本曲线穿过 0 的月份'))
$roiX.Add((RoiRow '结果' '续年投入产出比' (Cell '年收益 ÷ 年运行成本') (Fx '=D19/D5' '0.0"倍"') '' '首年之后每年只付年服务费' 'total'))
$roiX.Add((RoiRow '结果' '三年 ROI' (Cell '(3 年收益 − 3 年总成本) ÷ 3 年总成本') (Fx '=(3*D19-(D4+3*D5))/(D4+3*D5)' '0.0%') '' '' 'total'))
$roiX.Add((Row 'blank' @()))
$roiX.Add((Row 'sub' @((Cell '敏感性分析'), (Cell '情景'), (Cell '计算式'), (Cell '数值'), (Cell '单位'), (Cell '说明'))))
$scen = @(
    @{ n = '保守'; f = ('(1-{0})' -f (PN 'sens')) }
    @{ n = '基准'; f = '1' }
    @{ n = '乐观'; f = ('(1+{0})' -f (PN 'sens')) }
)
foreach ($s in $scen) {
    $roiX.Add((RoiRow ('敏感性 · ' + $s.n) '首年 ROI' (Cell '（年收益 × 系数 − 首年总投入）÷ 首年总投入') (Fx ('=(D19*{0}-D6)/D6' -f $s.f) '0.0%') '' ''))
    $roiX.Add((RoiRow ('敏感性 · ' + $s.n) '回本周期（课程公式）' (Cell '首年总投入 ÷（年收益 × 系数 ÷ 12）') (Fx ('=D6/(D19*{0}/12)' -f $s.f) '0.0') '月' ''))
    $roiX.Add((RoiRow ('敏感性 · ' + $s.n) '三年 ROI' (Cell '（3 年收益 × 系数 − 3 年总成本）÷ 3 年总成本') (Fx ('=(3*D19*{0}-(D4+3*D5))/(D4+3*D5)' -f $s.f) '0.0%') '' ''))
}
$roiX.Add((RoiRow '敏感性' '仅工时口径（不计避免冲突）：年净收益' (Cell '协调工时收益 − 年运行成本') (Fx '=D12-D5' '#,##0') '元/年' '为负、不回本——与阶段 04 的结论一致：只算工时撑不起这个报价'))
$roiX.Add((Row 'blank' @()))
$roiX.Add((Row 'sub' @((Cell '回本曲线'), (Cell '时点'), (Cell '计算式'), (Cell '数值'), (Cell '单位'), (Cell '说明'))))
foreach ($m in @(0, 6, 12, 18, 24, 30, 36)) {
    $roiX.Add((RoiRow '回本曲线' ('第 ' + $m + ' 月末累计净现金流') (Cell '− 一次性投入 + 月净收益 × 月数') (Fx ('=-$D$4+($D$19-$D$5)/12*{0}' -f $m) '#,##0') '元' '基准情景'))
}
$roiX.Add((Row 'blank' @()))
$roiX.Add((Row 'sub' @((Cell '首年回本的前提'), (Cell '项目'), (Cell '计算式'), (Cell '数值'), (Cell '单位'), (Cell '对照当前假设'))))
$roiX.Add((RoiRow '回本前提' '首年回本所需年收益' (Cell '首年总投入') (Fx '=D6' '#,##0') '元/年' '业务量够大，或冲突损失够大，二者有其一即可首年回本'))
function RoiRowS([string]$a, [string]$b, [object]$calc, [object]$val, [string]$unit, [string]$srcFormula) {
    return (Row 'data' @((Cell $a), (Cell $b), $calc, $val, (Cell $unit), (Fx $srcFormula)))
}
$roiX.Add((RoiRowS '回本前提' '仅靠工时：每月需节省工时' (Cell '首年回本所需年收益 ÷ 12 ÷（综合时薪 × 转化系数）') (Fx ('=D49/12/(D11*{0})' -f (PN 'conv')) '#,##0') '小时/月' '="现假设 "&TEXT(D9,"0.0")&" 小时/月"'))
$roiX.Add((RoiRowS '回本前提' '仅靠工时：对应月单量' (Cell '每月需节省工时 ÷（每单协调工时 × 压降比例）') (Fx '=D50/(D7*D8)' '#,##0') '单/月' ('="现假设 "&TEXT({0},"0")&" 单/月；客服日均 45 条咨询若多数是改单，真实单量可能远高于此"' -f (PN 'reports'))))
$roiX.Add((RoiRowS '回本前提' '仅靠避免冲突：对应单次冲突损失' (Cell '首年回本所需年收益 ÷（年冲突次数 × 降低比例）') (Fx '=D49/(D14*D16)' '#,##0') '元/次' '="现假设 "&TEXT(D15,"#,##0")&" 元/次"'))
$roiX.Add((RoiRowS '回本前提' '仅靠避免冲突：对应年冲突次数' (Cell '首年回本所需年收益 ÷（单次冲突损失 × 降低比例）') (Fx '=D49/(D15*D16)' '0.0') '次/年' '="现假设 "&TEXT(D14,"0")&" 次/年"'))
$roiX.Add((Row 'blank' @()))
$roiX.Add((Row 'sub' @((Cell '收益构成'), (Cell '项目'), (Cell '计算式'), (Cell '数值'), (Cell '单位'), (Cell '说明'))))
$roiX.Add((RoiRow '收益构成' '协调工时收益' (Cell '见上') (Fx '=D12' '#,##0') '元/年' ''))
$roiX.Add((RoiRow '收益构成' '避免冲突收益' (Cell '见上') (Fx '=D17' '#,##0') '元/年' ''))
$wsRoi = Add-Sheet $xl '5-客户ROI一页纸' $roiX @(18, 34, 44, 16, 9, 48) 3
try {
    # 图表用的短标签放在 H 列
    $months = @(0, 6, 12, 18, 24, 30, 36)
    $wsRoi.Cells[39, 8].Value = '月份'
    for ($k = 0; $k -lt $months.Count; $k++) { $wsRoi.Cells[(40 + $k), 8].Value = ('第 ' + $months[$k] + ' 月') }
    $wsRoi.Column(8).Width = 9
    $wsRoi.Cells['H39:H46'].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $wsRoi.Cells['H39:H46'].Style.Font.Color.SetColor($C_GREY)
    $wsRoi.Cells['H39:H46'].Style.Font.Size = 9
    $topPx = 0.0
    for ($r = 1; $r -le 57; $r++) { $h = [double](@($wsRoi.Row($r).Height)[0]); if ($h -le 0) { $h = 15 }; $topPx += $h * 4 / 3 }
    $topPx = [int]$topPx + 24
    $c1 = $wsRoi.Drawings.AddChart('CumCash', [OfficeOpenXml.Drawing.Chart.eChartType]::Line)
    $ser1 = $c1.Series.Add($wsRoi.Cells['D40:D46'], $wsRoi.Cells['H40:H46'])
    $c1.Title.Text = '累计净现金流（基准情景，元）'
    $c1.SetPosition($topPx, 0); $c1.SetSize(640, 300)
    $c1.Legend.Remove()
    try { $ser1.Marker = [OfficeOpenXml.Drawing.Chart.eMarkerStyle]::Circle; $ser1.DataLabels.ShowValue = $true; $ser1.DataLabels.NumberFormat.Format = '#,##0' } catch { }
    $c2 = $wsRoi.Drawings.AddChart('GainMix', [OfficeOpenXml.Drawing.Chart.eChartType]::BarClustered)
    $ser2 = $c2.Series.Add($wsRoi.Cells['D56:D57'], $wsRoi.Cells['B56:B57'])
    $c2.Title.Text = '年收益构成（元）'
    $c2.SetPosition($topPx, 700); $c2.SetSize(520, 300)
    $c2.Legend.Remove()
    try {
        $c2.YAxis.MinorTickMark = 'None'
        $ser2.DataLabels.ShowValue = $true; $ser2.DataLabels.NumberFormat.Format = '#,##0'
    } catch { }
} catch { Write-Host ('ROI 图表未生成：' + $_.Exception.Message + ' @line ' + $_.InvocationInfo.ScriptLineNumber) }

# ---- 4-部署方案对比表（表头 3、三方案 4-6、盈亏平衡 7、测算 10-14）
$depX = [System.Collections.Generic.List[object]]::new()
$depX.Add((Row 'title' @((Cell '部署方案对比表'))))
$depX.Add((Row 'note' @((Cell '核心核对在任何方案下都在内网由规则引擎完成，数据不出内网；三方案的区别只在“报告说明文字怎么生成”。云端 API、本地自建、混合三种都算了一遍账。'))))
$depX.Add((Row 'header' @((Cell '方案'), (Cell '说明文字怎么生成'), (Cell '数据是否出内网'), (Cell '业务假设'), (Cell '成本构成'), (Cell '年成本(元)'), (Cell '三年成本(元)'), (Cell '优点'), (Cell '不选的原因'))))
$depX.Add((Row 'sub' @((Cell '一、云端 API（推荐）'), (Cell '全部报告调 DeepSeek Flash 生成'), (Cell '脱敏后出内网（待确认）'), (Cell '月 40 份报告；上下文约 4.4K Token；无并发压力'), (Cell '纯 Token 费，无固定成本'), (Fx ('={0}' -f $T_YEAR) '0.00'), (Fx '=F4*3' '0.00'), (Cell '固定成本为零，零运维，模型质量随厂商升级'), (Cell '依赖“脱敏后可出内网”这一前提，须信息化负责人书面确认'))))
$depX.Add((Row 'data' @((Cell '二、本地自建'), (Cell '内网部署 7B 开源模型（INT8）生成'), (Cell '不出内网'), (Cell '峰值并发 2 路；上下文 8K；单卡 RTX 4090 24G'), (Cell '自购工作站按 3 年折旧 + 年电费 + 运维；部署 5 人天'), (Fx ('={0}/{1}+{2}*24*365*{3}*{4}+{5}*12' -f (PN 'gpu'), (PN 'depYears'), (PN 'kw'), (PN 'elec'), (PN 'pue'), (PN 'modelOps')) '#,##0.00'), (Fx ('={0}+3*({1}*24*365*{2}*{3}+{4}*12)+{5}*{6}G{7}' -f (PN 'gpu'), (PN 'kw'), (PN 'elec'), (PN 'pue'), (PN 'modelOps'), (PN 'deployDays'), $S2b, $fdeRow) '#,##0.00'), (Cell '数据不出内网，断网可用'), (Fx ('="年固定成本约为 API 方案的 "&TEXT(F5/F4,"#,##0")&" 倍，而实际月 Token 量仅为盈亏平衡点的 "&TEXT(B14,"0.00%")')))))
$depX.Add((Row 'data' @((Cell '三、混合'), (Cell '固定模板为底，仅“冲突 / 需审批”类报告调 API'), (Fx ('="仅 "&TEXT({0},"0%")&" 脱敏后出内网"' -f (PN 'mix'))), (Cell '同方案一'), (Fx ('="API 费用 × "&TEXT({0},"0%")' -f (PN 'mix'))), (Fx ('=F4*{0}' -f (PN 'mix')) '0.00'), (Fx '=F6*3' '0.00'), (Cell '成本最低，API 故障时仍能出报告'), (Cell '多数报告的可读性不如模型生成；适合作为方案一未获确认时的过渡'))))
$depX.Add((Row 'total' @((Cell '盈亏平衡点'), (Fx '="自建月固定成本 "&TEXT(B10,"#,##0")&" 元 ÷ API 综合单价 "&TEXT(B11,"0.00")&" 元/百万 Token ≈ "&TEXT(B12/100000000,"0.0")&" 亿 Token/月；本项目实际用量仅为其 "&TEXT(B14,"0.00%")'), (Cell ''), (Cell '低于临界用量调 API 更便宜；高于才轮到自建，且还要先过合规与运维能力两关'), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''))))
$depX.Add((Row 'blank' @()))
$depX.Add((Row 'section' @((Cell '盈亏平衡测算'), (Cell ''), (Cell '测算式'), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''))))
$depX.Add((Row 'data' @((Cell '自建月固定成本（元/月）'), (Fx '=F5/12' '#,##0.00'), (Cell '自建年成本 ÷ 12'), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''))))
$depX.Add((Row 'data' @((Cell 'API 综合单价（元/百万 Token）'), (Fx ('=({0}*{1}+{2}*{3})/({0}+{2})' -f $T_IN, $T_PRICE_IN, $T_OUT, $T_PRICE_OUT) '0.00'), (Cell '（输入 Token × 输入单价 + 输出 Token × 输出单价）÷ 总 Token'), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''))))
$depX.Add((Row 'data' @((Cell '临界月 Token 量'), (Fx '=B10/B11*1000000' '#,##0'), (Cell '自建月固定成本 ÷ API 综合单价 × 1,000,000'), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''))))
$depX.Add((Row 'data' @((Cell '本项目实际月 Token 量'), (Fx ('={0}' -f $T_TOK) '#,##0'), (Cell '“3-Token月成本测算表”基准档'), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''))))
$depX.Add((Row 'data' @((Cell '实际用量 ÷ 临界用量'), (Fx '=B13/B12' '0.00%'), (Cell '远小于 1，说明自建不划算'), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''), (Cell ''))))
Add-Sheet $xl '4-部署方案对比表' $depX @(24, 40, 20, 34, 38, 14, 14, 30, 44) 3 | Out-Null

# ---- 4-推荐配置单
$t4b = Convert-Table $Tables['04-推荐配置单.csv'] @{}
$row = $t4b[5]
$row.c[2] = Fx ('="权重 "&TEXT({0},"0")&" GB（7B × INT8 1 字节）+ KV Cache "&TEXT({1}*{2}*{3}/1000000,"0.00")&" GB（"&TEXT({1},"0")&" KB × "&TEXT({2},"#,##0")&" × "&TEXT({3},"0")&" 路）= "&TEXT({0}+{1}*{2}*{3}/1000000,"0.00")&" GB，× "&TEXT({4},"0.00")&" 开销 ≈ "&TEXT(({0}+{1}*{2}*{3}/1000000)*{4},"0.00")&" GB，单卡 24G 可承载"' -f (PN 'weightGb'), (PN 'kvKb'), (PN 'ctx'), (PN 'conc'), (PN 'overhead'))
$row.c[1] = Fx ('="单卡 RTX 4090 24G ×1（显存需求约 "&TEXT(({0}+{1}*{2}*{3}/1000000)*{4},"0.0")&" GB）；系统内存 32G；存储 1TB SSD"' -f (PN 'weightGb'), (PN 'kvKb'), (PN 'ctx'), (PN 'conc'), (PN 'overhead'))
$row = $t4b[7]
$row.c[1] = Fx ('="月均 "&TEXT({0},"0")&" 次请求、约 "&TEXT({1},"#,##0")&" Token 内，包干含 "&TEXT({2},"0%")&" 缓冲"' -f (PN 'reports'), $T_TOK, (PN 'packBuf'))
$rows4b = With-Header '推荐配置单' '对应“4-部署方案对比表”的推荐方案（云端 API），并列出数据不可出内网时的备选配置。' $t4b
Add-Sheet $xl '4-推荐配置单' $rows4b @(24, 62, 80) 3 | Out-Null

# ---- 0-项目一页纸
$rows0 = With-Header '项目一页纸' '开课前先定的五个关键数。华衡为合成业务现场，标“假设”的值为测算设定值。' (Convert-Table $Tables['00-项目一页纸.csv'] @{})
Add-Sheet $xl '0-项目一页纸' $rows0 @(18, 78, 52) 3 | Out-Null

# ---- 参数
$parX = [System.Collections.Generic.List[object]]::new()
$parX.Add((Row 'title' @((Cell '参数与假设'))))
$parX.Add((Row 'note' @((Cell '改这里的取值，其余各工作表的公式会联动重算。带“假设”的项在案例中没有记录，签约前须由客户提供基线数据后替换。'))))
$parX.Add((Row 'header' @((Cell '参数'), (Cell '取值'), (Cell '说明'))))
foreach ($d in $paramDefs) { $parX.Add((Row 'data' @((Cell $d.n), (Cell $d.v $d.fmt), (Cell $d.s)))) }
Add-Sheet $xl '参数' $parX @(34, 16, 64) 3 | Out-Null

# ---- 说明（放到最前）
$smX = [System.Collections.Generic.List[object]]::new()
$smX.Add((Row 'title' @((Cell '华衡智造 · 订单变更影响核对助手 —— 企业报价包'))))
$smX.Add((Row 'note' @((Cell '按《FDE 企业报价与 ROI 测算》课件口径制作。各工作表的汇总格都是公式，输入集中在“参数”表；蓝色表头为数据区，浅黄底为对外报价。华衡是合成业务现场，标“假设”的输入不是事实。'))))
$smX.Add((Row 'header' @((Cell '关键结果'), (Cell '数值'), (Cell '单位'), (Cell '出处'))))
$kpis = @(
    @{ n = '人天合计（含缓冲）'; f = ('={0}E{1}' -f $S2, $rTot); fmt = '#,##0'; u = '人天'; s = '2-人天估算与单价表' }
    @{ n = '人力成本（成本口径）'; f = ('={0}H{1}' -f $S2, $rTot); fmt = '#,##0'; u = '元'; s = '2-人天估算与单价表' }
    @{ n = '模型费（基准档年成本）'; f = ('={0}' -f $T_YEAR); fmt = '0.00'; u = '元/年'; s = '3-Token月成本测算表' }
    @{ n = '一次性实施费'; f = ('={0}G35' -f $S1); fmt = '#,##0'; u = '元'; s = '1-成本拆解表 / 6-报价汇总表' }
    @{ n = '年服务费'; f = ('={0}G36' -f $S1); fmt = '#,##0'; u = '元/年'; s = '1-成本拆解表 / 6-报价汇总表' }
    @{ n = '首年合计'; f = ('={0}G37' -f $S1); fmt = '#,##0'; u = '元'; s = '6-报价汇总表' }
    @{ n = '续年合计'; f = ('={0}G36' -f $S1); fmt = '#,##0'; u = '元/年'; s = '6-报价汇总表' }
    @{ n = '年收益合计'; f = ('={0}D19' -f $S5); fmt = '#,##0'; u = '元/年'; s = '5-客户ROI一页纸' }
    @{ n = '回本周期（课程公式）'; f = ('={0}D22' -f $S5); fmt = '0.0'; u = '月'; s = '5-客户ROI一页纸' }
    @{ n = '续年投入产出比'; f = ('={0}D24' -f $S5); fmt = '0.0"倍"'; u = ''; s = '5-客户ROI一页纸' }
    @{ n = '三年 ROI'; f = ('={0}D25' -f $S5); fmt = '0.0%'; u = ''; s = '5-客户ROI一页纸' }
    @{ n = '首年 ROI（含一次性实施费）'; f = ('={0}D21' -f $S5); fmt = '0.0%'; u = ''; s = '5-客户ROI一页纸' }
)
foreach ($k in $kpis) { $smX.Add((Row 'data' @((Cell $k.n), (Fx $k.f $k.fmt), (Cell $k.u), (Cell $k.s)))) }
$smX.Add((Row 'blank' @()))
$smX.Add((Row 'header' @((Cell '待客户确认的假设'), (Cell '取值'), (Cell '影响'), (Cell '向谁确认'))))
$asm = @(
    @{ n = '月核对报告量'; f = ('=TEXT({0},"0")&" 份/月"' -f (PN 'reports')); i = 'Token 量、人力节省'; w = '生产计划经理（试点期实测）' }
    @{ n = '岗位人数与月薪'; f = ('=TEXT({0},"0")&" 人、月薪 "&TEXT({1},"#,##0")&" 元"' -f (PN 'people'), (PN 'salary')); i = 'ROI 人力节省'; w = '人事' }
    @{ n = '年冲突次数与单次损失'; f = ('=TEXT({0},"0")&" 次 × "&TEXT({1},"#,##0")&" 元"' -f (PN 'errN'), (PN 'errLoss')); i = 'ROI 的决定性输入'; w = '销售总监、生产计划经理（历史记录）' }
    @{ n = '每单协调工时、压降比例'; f = ('=TEXT({0},"0.00")&" 小时/单、压降 "&TEXT({1},"0%")' -f (PN 'hoursOrder'), (PN 'cut')); i = 'ROI 协调工时收益'; w = '生产计划经理、客户服务主管（抽样计时）' }
    @{ n = '脱敏后数据可出内网'; f = '待确认'; i = '部署路线；不成立则回退固定模板或追加内网 7B 部署'; w = '信息化负责人' }
    @{ n = '开发测试云主机、差旅单价'; f = ('=TEXT({0},"#,##0")&" 元/月、"&TEXT({1},"#,##0")&" 元/次"' -f (PN 'devVm'), (PN 'travelOnce')); i = '实施费'; w = '我方' }
)
foreach ($a in $asm) {
    if ($a.f.StartsWith('=')) { $vc = Fx $a.f } else { $vc = Cell $a.f }
    $smX.Add((Row 'data' @((Cell $a.n), $vc, (Cell $a.i), (Cell $a.w))))
}
Add-Sheet $xl '说明' $smX @(30, 26, 30, 44) 3 | Out-Null
$xl.Workbook.Worksheets.MoveToStart('说明')

# ---------------------------------------------------------------- 计算、保存
try { $xl.Workbook.FullCalcOnLoad = $true } catch { }
Close-ExcelPackage $xl -Calculate
Write-Host ('已生成 ' + (Split-Path -Leaf $xlsx))

# ---------------------------------------------------------------- Markdown 版（GitHub 直接渲染）
function MdCell([object]$v) {
    $s = [string]$v
    if ($s -match '^-?\d+$' -and [math]::Abs([double]$s) -ge 1000) { $s = ([double]$s).ToString('N0', $inv) }
    return ($s -replace '\|', '\|')
}
function Md-Table([object[]]$t) {
    $sb = [System.Text.StringBuilder]::new()
    $hdr = @($t[0] | ForEach-Object { MdCell $_ })
    [void]$sb.AppendLine('| ' + ($hdr -join ' | ') + ' |')
    [void]$sb.AppendLine('|' + (($hdr | ForEach-Object { ' --- ' }) -join '|') + '|')
    for ($i = 1; $i -lt $t.Count; $i++) {
        $row = @($t[$i])
        if ($row.Count -eq 1) { [void]$sb.AppendLine(); [void]$sb.AppendLine('> ' + (MdCell $row[0])); continue }
        [void]$sb.AppendLine('| ' + ((@($row | ForEach-Object { MdCell $_ })) -join ' | ') + ' |')
    }
    return $sb.ToString()
}
$md = [System.Text.StringBuilder]::new()
[void]$md.AppendLine('# 华衡智造 · 订单变更影响核对助手 —— 企业报价包')
[void]$md.AppendLine()
[void]$md.AppendLine('按《FDE 企业报价与 ROI 测算》课件口径制作，数字由 [`tools/build_pricing.ps1`](tools/build_pricing.ps1) 一次算出，各表可以互相复算。带公式的 Excel 版见 [`华衡智造-企业报价包.xlsx`](华衡智造-企业报价包.xlsx)，口径、假设清单与数字咬合校验见 [`README.md`](README.md)。')
[void]$md.AppendLine()
[void]$md.AppendLine('> 华衡是合成业务现场，标“假设”的输入不是事实；签约前须由客户提供基线数据后重算。金额除注明外均为元。')
$sections = @(
    @{ n = '0'; t = '项目一页纸'; f = @('00-项目一页纸.csv') }
    @{ n = '1'; t = '成本拆解表（含人天估算与报价）'; f = @('01-成本拆解表.csv') }
    @{ n = '2'; t = '人天估算与单价表'; f = @('02-人天估算与单价表.csv', '02-人天单价推导.csv') }
    @{ n = '3'; t = 'Token 月成本测算表'; f = @('03-单次输入构成.csv', '03-Token月成本测算表.csv') }
    @{ n = '4'; t = '部署方案对比表与推荐配置单'; f = @('04-部署方案对比表.csv', '04-推荐配置单.csv') }
    @{ n = '5'; t = '客户 ROI 一页纸'; f = @('05-客户ROI一页纸.csv') }
    @{ n = '6'; t = '报价汇总表'; f = @('06-报价汇总表.csv') }
)
$subTitle = @{
    '02-人天估算与单价表.csv' = '工作包'; '02-人天单价推导.csv' = '单价推导'
    '03-单次输入构成.csv' = '单次输入构成'; '03-Token月成本测算表.csv' = '三档测算'
    '04-部署方案对比表.csv' = '三方案对比'; '04-推荐配置单.csv' = '推荐配置单'
}
foreach ($s in $sections) {
    [void]$md.AppendLine(); [void]$md.AppendLine('## ' + $s.n + ' ' + $s.t); [void]$md.AppendLine()
    foreach ($f in $s.f) {
        if ($s.f.Count -gt 1) { [void]$md.AppendLine('**' + $subTitle[$f] + '**'); [void]$md.AppendLine() }
        [void]$md.Append((Md-Table $Tables[$f]))
    }
    if ($s.n -eq '5') {
        $pts = @(foreach ($m in @(0, 6, 12, 18, 24, 30, 36)) { R0 (-$I1 + $roiBase.netMonth * $m) })
        $yMin = [math]::Floor((($pts | Measure-Object -Minimum).Minimum) / 50000) * 50000
        $yMax = [math]::Ceiling((($pts | Measure-Object -Maximum).Maximum) / 50000) * 50000
        [void]$md.AppendLine()
        [void]$md.AppendLine('```mermaid')
        [void]$md.AppendLine('pie showData')
        [void]$md.AppendLine('    title 年收益构成（基准，元）')
        [void]$md.AppendLine('    "协调工时" : ' + (R0 $gainLabor))
        [void]$md.AppendLine('    "避免冲突" : ' + (R0 $gainErr))
        [void]$md.AppendLine('```')
        [void]$md.AppendLine()
        [void]$md.AppendLine('```mermaid')
        [void]$md.AppendLine('xychart-beta')
        [void]$md.AppendLine('    title "累计净现金流（基准情景，元）"')
        [void]$md.AppendLine('    x-axis ["第0月", "第6月", "第12月", "第18月", "第24月", "第30月", "第36月"]')
        [void]$md.AppendLine(('    y-axis "元" {0} --> {1}' -f $yMin, $yMax))
        [void]$md.AppendLine('    line [' + ($pts -join ', ') + ']')
        [void]$md.AppendLine('```')
    }
}
[System.IO.File]::WriteAllText((Join-Path $OutDir '报价包.md'), $md.ToString(), [System.Text.UTF8Encoding]::new($false))
Write-Host '已生成 报价包.md'
