# =====================================================================
# 公式助手 - 自动化冒烟测试
# 通过 Excel COM 加载 dist\公式助手.xlam，调用其中的公开函数做断言。
# 用法：powershell -ExecutionPolicy Bypass -File build\test.ps1
# =====================================================================
param([string]$XlamPath)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'excel-common.ps1')
if (-not $XlamPath) { $XlamPath = Join-Path $root 'dist\公式助手.xlam' }
if (-not (Test-Path -LiteralPath $XlamPath)) { Write-Error "未找到 $XlamPath ，请先构建" }

$script:failures = 0
function Assert-Eq($name, $expected, $actual) {
    if ("$expected" -eq "$actual") {
        Write-Host ("PASS  " + $name)
    } else {
        $script:failures++
        Write-Host ("FAIL  " + $name)
        Write-Host ("      expected: " + $expected)
        Write-Host ("      actual  : " + $actual)
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    [void]$excel.Workbooks.Open($XlamPath)

    # ---------- 纯函数：VLOOKUP ----------
    $r = $excel.Run('BuildVlookupFormula', '$B$2', '$D$2:$E$10', '2', $true, $true)
    Assert-Eq 'VLOOKUP 精确+IFERROR' '=IFERROR(VLOOKUP($B$2,$D$2:$E$10,2,FALSE),"")' $r
    $r = $excel.Run('BuildVlookupFormula', '"香蕉"', '$D$2:$E$10', '3', $false, $false)
    Assert-Eq 'VLOOKUP 文本+模糊' '=VLOOKUP("香蕉",$D$2:$E$10,3,TRUE)' $r

    # ---------- 纯函数：参数解析 ----------
    $r = $excel.Run('FormatArg', '香蕉')
    Assert-Eq 'FormatArg 文本加引号' '"香蕉"' $r
    $r = $excel.Run('FormatArg', '$B$2')
    Assert-Eq 'FormatArg 引用原样' '$B$2' $r
    $r = $excel.Run('ParseCriteria', '>=60')
    Assert-Eq 'ParseCriteria 比较条件加引号' '">=60"' $r
    $r = $excel.Run('ParseCriteria', '完成')
    Assert-Eq 'ParseCriteria 文本加引号' '"完成"' $r
    $r = $excel.Run('ParseCriteria', '张*')
    Assert-Eq 'ParseCriteria 通配符' '"张*"' $r
    $r = $excel.Run('ParseCriteria', '$D$2')
    Assert-Eq 'ParseCriteria 引用原样' '$D$2' $r
    $r = $excel.Run('ParseCriteria', '>B2')
    Assert-Eq 'ParseCriteria 与单元格比较' '">"&B2' $r
    $r = $excel.Run('IsCellAddress', 'A1:B2')
    Assert-Eq 'IsCellAddress 区域' 'True' $r

    # ---------- 纯函数：通用模板构建 ----------
    $tok = [object[]]@('$D$2:$D$31', '">=60"')
    $r = $excel.Run('BuildFormulaFromTemplate', 'COUNTIF', '', $tok)
    Assert-Eq 'COUNTIF 构建' '=COUNTIF($D$2:$D$31,">=60")' $r

    $tok = [object[]]@('$E$2:$E$31', '$D$2:$D$31', '"苹果"', '$F$2:$F$31', '">=10"', '', '')
    $r = $excel.Run('BuildFormulaFromTemplate', 'SUMIFS', '', $tok)
    Assert-Eq 'SUMIFS 构建（可选参数跳过）' '=SUMIFS($E$2:$E$31,$D$2:$D$31,"苹果",$F$2:$F$31,">=10")' $r

    $tok = [object[]]@('$E$2:$E$31', '$G$1', '$D$2:$D$31')
    $r = $excel.Run('BuildFormulaFromTemplate', 'INDEXMATCH', 'INDEX({1},MATCH({2},{3},0))', $tok)
    Assert-Eq 'INDEX+MATCH 模板' '=INDEX($E$2:$E$31,MATCH($G$1,$D$2:$D$31,0))' $r

    $tok = [object[]]@('$E$2:$E$31', '$D$2:$D$31', '', '$F$2:$F$31', '">=10"', '', '')
    $r = $excel.Run('BuildFormulaFromTemplate', 'SUMIFS', '', $tok)
    Assert-Eq 'SUMIFS 空参数跳跃报错' '' $r

    # ---------- 纯函数：包裹 ----------
    $r = $excel.Run('WrapFormulaText', '=SUM(A1:A9)', '""')
    Assert-Eq 'IFERROR 包裹' '=IFERROR(SUM(A1:A9),"")' $r
    $r = $excel.Run('FormatLiteralToken', '')
    Assert-Eq '容错值空 -> 空文本' '""' $r
    $r = $excel.Run('FormatLiteralToken', '0')
    Assert-Eq '容错值数字原样' '0' $r

    # ---------- 配置表 ----------
    $r = $excel.Run('HasFuncDef', 'COUNTIF')
    Assert-Eq 'HasFuncDef COUNTIF 存在' 'True' $r
    $r = $excel.Run('HasFuncDef', 'INDEXMATCH')
    Assert-Eq 'HasFuncDef INDEXMATCH 存在' 'True' $r
    $r = $excel.Run('HasFuncDef', 'NO_SUCH_FUNC')
    Assert-Eq 'HasFuncDef 不存在的函数' 'False' $r

    # ---------- 工作表集成测试 ----------
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Activate()
    $ws.Range('B1').Value2 = 'banana'
    $ws.Range('D2').Value2 = 'apple';  $ws.Range('E2').Value2 = 10
    $ws.Range('D3').Value2 = 'banana'; $ws.Range('E3').Value2 = 20
    $ws.Range('D4').Value2 = 'banana'; $ws.Range('E4').Value2 = 30

    $f = $excel.Run('BuildVlookupFormula', '$B$1', '$D$2:$E$4', '2', $true, $true)
    $ws.Range('A1').Formula = $f
    Assert-Eq 'VLOOKUP 实算结果' '20' "$($ws.Range('A1').Value2)"

    $tok = [object[]]@('$E$2:$E$4', '">=20"')
    $f = $excel.Run('BuildFormulaFromTemplate', 'COUNTIF', '', $tok)
    $ws.Range('A2').Formula = $f
    Assert-Eq 'COUNTIF 实算结果' '2' "$($ws.Range('A2').Value2)"

    $tok = [object[]]@('$D$2:$D$4', '"banana"', '$E$2:$E$4')
    $f = $excel.Run('BuildFormulaFromTemplate', 'SUMIF', '', $tok)
    $ws.Range('A3').Formula = $f
    Assert-Eq 'SUMIF 实算结果' '50' "$($ws.Range('A3').Value2)"

    # IFERROR 包裹集成（核心函数无对话框，可自动化）
    $ws.Range('C1').Formula = '=SUM(E2:E4)'
    $ws.Range('C1').Select()
    $n = $excel.Run('WrapIfErrorCore', '""')
    Assert-Eq 'IFERROR 包裹数量' '1' "$n"
    Assert-Eq 'IFERROR 包裹后公式' '=IFERROR(SUM(E2:E4),"")' $ws.Range('C1').Formula
    Assert-Eq 'IFERROR 包裹后结果' '60' "$($ws.Range('C1').Value2)"

    # ROUND 包裹集成
    $ws.Range('C2').Formula = '=E2/E4'
    $ws.Range('C2').Select()
    $n = $excel.Run('WrapRoundCore', 2)
    Assert-Eq 'ROUND 包裹数量' '1' "$n"
    Assert-Eq 'ROUND 包裹后公式' '=ROUND(E2/E4,2)' $ws.Range('C2').Formula
    Assert-Eq 'ROUND 包裹后结果' '0.33' "$($ws.Range('C2').Value2)"

    # TEXT 包裹集成
    $n = $excel.Run('WrapTextCore', '"0.00"')
    Assert-Eq 'TEXT 包裹数量' '1' "$n"
    Assert-Eq 'TEXT 包裹后公式' '=TEXT(ROUND(E2/E4,2),"0.00")' $ws.Range('C2').Formula
    Assert-Eq 'TEXT 包裹后结果' '0.33' "$($ws.Range('C2').Value2)"

    # InsertFormulaAt
    $r = $excel.Run('InsertFormulaAt', $ws.Range('A5'), '=1+1', $true)
    Assert-Eq 'InsertFormulaAt 成功' 'True' "$r"
    Assert-Eq 'InsertFormulaAt 公式' '=1+1' $ws.Range('A5').Formula

    # ---------- 引用方式（$ 判定）：需要活动工作表上下文，放在建簿之后 ----------
    $r = $excel.Run('QualifiedAddressFromText', 'D1:E10', 'relative')
    Assert-Eq 'QualifiedAddress relative' 'D1:E10' $r
    $r = $excel.Run('QualifiedAddressFromText', 'D1:E10', 'absolute')
    Assert-Eq 'QualifiedAddress absolute' '$D$1:$E$10' $r
    $r = $excel.Run('QualifiedAddressFromText', 'D1:E10', 'rowabs')
    Assert-Eq 'QualifiedAddress rowabs（锁行）' 'D$1:E$10' $r
    $r = $excel.Run('QualifiedAddressFromText', 'D1:E10', 'colabs')
    Assert-Eq 'QualifiedAddress colabs（锁列）' '$D1:$E10' $r
    $r = $excel.Run('QualifiedAddressFromText', 'D1:E10', '')
    Assert-Eq 'QualifiedAddress auto 多格' '$D$1:$E$10' $r
    $r = $excel.Run('QualifiedAddressFromText', 'B1', '')
    Assert-Eq 'QualifiedAddress auto 单格' 'B1' $r
    $r = $excel.Run('QualifiedAddressFromText', 'B1', 'absolute')
    Assert-Eq 'QualifiedAddress 单格绝对' '$B$1' $r
    $r = $excel.Run('QualifiedAddressFromText', 'B1:C2', 'xyz')
    Assert-Eq 'QualifiedAddress 非法模式兜底 auto' '$B$1:$C$2' $r
    $r = $excel.Run('NormalizeRefMode', '绝对')
    Assert-Eq 'NormalizeRefMode 中文' 'absolute' $r

    $wb.Close($false)

    # ---------- 窗体级端到端测试（需要 -IncludeTests 构建的测试版） ----------
    $testXlam = Join-Path $root 'dist\公式助手-test.xlam'
    if (Test-Path -LiteralPath $testXlam) {
        Write-Host ''
        Write-Host '--- 窗体级端到端测试（测试构建）---'
        $excel2 = New-Object -ComObject Excel.Application
        $excel2.Visible = $false
        $excel2.DisplayAlerts = $false
        try {
            $twb = $excel2.Workbooks.Open($testXlam)
            $log = $excel2.Run('RunFormTests')
            foreach ($line in ("$log" -split "`n")) {
                if ($line.Trim().Length -gt 0) { Write-Host $line.TrimEnd() }
            }
            if ("$log" -notmatch 'ALL PASS') { $script:failures++ }
            $twb.Close($false)
        }
        catch {
            $script:failures++
            Write-Host ("FAIL  窗体测试执行异常: " + $_.Exception.Message)
        }
        finally {
            Stop-ExcelGracefully $excel2 '窗体测试 Excel'
        }
    } else {
        Write-Host ''
        Write-Host '（跳过窗体级测试：未找到 dist\公式助手-test.xlam，用 build.ps1 -IncludeTests 生成）'
    }
}
finally {
    Stop-ExcelGracefully $excel '测试 Excel'
}

Write-Host ''
if ($script:failures -eq 0) {
    Write-Host '== 全部测试通过 =='
    exit 0
} else {
    Write-Host ("== 有 " + $script:failures + " 个断言失败 ==")
    exit 1
}
