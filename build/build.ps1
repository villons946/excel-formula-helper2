param([switch]$IncludeTests)

# =====================================================================
# 公式助手 - 构建脚本
# 从纯文本源码（src\vba\*.bas + *.code.vba + 窗体布局）构建 dist\公式助手.xlam
# 用法：powershell -ExecutionPolicy Bypass -File build\build.ps1 [-IncludeTests]
#   -IncludeTests 额外生成 dist\公式助手-test.xlam（含窗体级自测模块）
# 说明：构建需要临时开启"信任对 VBA 工程对象模型的访问"（结束后自动还原），
#       向导窗体由本脚本通过 Excel 动态创建，仓库中不保存二进制 .frx。
# =====================================================================
$ErrorActionPreference = 'Stop'
$root      = Split-Path -Parent $PSScriptRoot
$vbaDir    = Join-Path $root 'src\vba'
$testDir   = Join-Path $root 'src\tests'
$ribbonXml = Join-Path $root 'src\ribbon\customui.xml'
$distDir   = Join-Path $root 'dist'
$outName   = if ($IncludeTests) { '公式助手-test.xlam' } else { '公式助手.xlam' }
$outXlam   = Join-Path $distDir $outName
$workDir   = Join-Path $PSScriptRoot '_work'

Write-Host '== 公式助手构建 =='

# ---------- 0. 临时开启 VBOM 访问（构建后还原） ----------
$secKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $secKey)) { New-Item -Path $secKey -Force | Out-Null }
$oldVBOM = $null
try { $oldVBOM = (Get-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction Stop).AccessVBOM } catch {}
Set-ItemProperty -Path $secKey -Name AccessVBOM -Value 1 -Type DWord

function Restore-VBOM {
    if ($null -eq $oldVBOM) {
        Remove-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $secKey -Name AccessVBOM -Value $oldVBOM -Type DWord
    }
}

# ---------- 1. 窗体布局 ----------
. (Join-Path $PSScriptRoot 'excel-common.ps1')
. (Join-Path $PSScriptRoot 'form-layouts.ps1')

function Add-CodeComponent($vbp, $name, $type, $code) {
    $comp = $vbp.VBComponents.Add($type)
    $comp.Name = $name
    [void]$comp.CodeModule.AddFromString($code)
}

# ---------- 2. Excel COM 构建 ----------
if (Test-Path -LiteralPath $outXlam) { Remove-Item -LiteralPath $outXlam -Force }
New-Item -ItemType Directory -Force -Path $distDir | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb  = $excel.Workbooks.Add()
    $vbp = $wb.VBProject

    # 标准模块 / 类模块
    Get-ChildItem -Path $vbaDir -Filter '*.bas' | ForEach-Object {
        Add-CodeComponent $vbp ([IO.Path]::GetFileNameWithoutExtension($_.Name)) 1 `
            ([IO.File]::ReadAllText($_.FullName, [Text.Encoding]::UTF8))
    }
    Get-ChildItem -Path $vbaDir -Filter '*.cls' | ForEach-Object {
        Add-CodeComponent $vbp ([IO.Path]::GetFileNameWithoutExtension($_.Name)) 2 `
            ([IO.File]::ReadAllText($_.FullName, [Text.Encoding]::UTF8))
    }
    if ($IncludeTests -and (Test-Path $testDir)) {
        Get-ChildItem -Path $testDir -Filter '*.bas' | ForEach-Object {
            Add-CodeComponent $vbp ([IO.Path]::GetFileNameWithoutExtension($_.Name)) 1 `
                ([IO.File]::ReadAllText($_.FullName, [Text.Encoding]::UTF8))
            Write-Host ("  已注入测试模块: " + $_.Name)
        }
    }

    # 动态创建向导窗体（vbext_ct_MSForm = 3）
    foreach ($formName in $FormLayouts.Keys) {
        $spec = $FormLayouts[$formName]
        $comp = $vbp.VBComponents.Add(3)
        $comp.Name = $formName
        $comp.Properties.Item('Caption').Value = $spec.Caption
        $comp.Properties.Item('Width').Value  = [single]$spec.Width
        $comp.Properties.Item('Height').Value = [single]$spec.Height
        $designer = $comp.Designer
        foreach ($c in $spec.Controls) {
            $ctl = $designer.Controls.Add($c.T, $c.N, $true)
            $ctl.Left   = [single]$c.L
            $ctl.Top    = [single]$c.Tp
            $ctl.Width  = [single]$c.W
            $ctl.Height = [single]$c.H
            if ($c.C)   { $ctl.Caption = $c.C }
            if ($c.ML)  { $ctl.MultiLine = $true; $ctl.ScrollBars = 2; $ctl.WordWrap = $true }
            if ($c.V)   { $ctl.Value = $true }
        }
        $codeFile = Join-Path $vbaDir ($formName + '.code.vba')
        [void]$comp.CodeModule.AddFromString([IO.File]::ReadAllText($codeFile, [Text.Encoding]::UTF8))
    }

    # 保存为加载宏（55 = xlAddIn）
    $wb.IsAddIn = $true
    $wb.SaveAs($outXlam, 55)
    $wb.Close($false)
    Write-Host "已生成 $outXlam"
}
finally {
    Stop-ExcelGracefully $excel '构建 Excel'
    Restore-VBOM
}

# ---------- 3. 注入 Ribbon XML ----------
function Inject-Ribbon($xlamPath, $xmlPath) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zipWork = Join-Path $workDir 'zip'
    if (Test-Path -LiteralPath $zipWork) { Remove-Item -LiteralPath $zipWork -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $zipWork | Out-Null

    $tmpZip = Join-Path $workDir 'xlam.zip'
    Copy-Item -LiteralPath $xlamPath -Destination $tmpZip -Force
    [IO.Compression.ZipFile]::ExtractToDirectory($tmpZip, $zipWork)

    # customUI 部件：与 Office 自带加载宏（SOLVER.XLAM / EUROTOOL.XLAM / FUNCRES.XLAM）保持一致
    #   - 部件路径 customui/customui.xml
    #   - XML 命名空间 http://schemas.microsoft.com/office/2006/01/customui
    #   - 关系挂在【包级】关系表 _rels/.rels，类型 2006/relationships/ui/extensibility
    # 注意：把关系写进 xl/_rels/workbook.xml.rels（或混用 2007/2009 命名空间）会导致
    #       Excel 静默忽略整个 Ribbon 自定义。
    New-Item -ItemType Directory -Force -Path (Join-Path $zipWork 'customui') | Out-Null
    Copy-Item -LiteralPath $xmlPath -Destination (Join-Path $zipWork 'customui\customui.xml') -Force

    $relsPath = Join-Path $zipWork '_rels\.rels'
    $rels = [IO.File]::ReadAllText($relsPath)
    if ($rels -notmatch 'extensibility') {
        $rels = $rels -replace '</Relationships>', '<Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customui/customui.xml" /></Relationships>'
        [IO.File]::WriteAllText($relsPath, $rels, (New-Object Text.UTF8Encoding($false)))
    }

    # 重新打包：必须逐条写入并强制正斜杠条目名
    # （ZipFile.CreateFromDirectory 在 .NET Framework 下用反斜杠，违反 OPC 规范，
    #   会导致 Excel 找不到 relationship 指向的 Ribbon 部件）
    $outZip = Join-Path $workDir 'out.xlam'
    if (Test-Path -LiteralPath $outZip) { Remove-Item -LiteralPath $outZip -Force }
    $baseLen = (Resolve-Path -LiteralPath $zipWork).Path.Length + 1
    $fs = [IO.File]::Open($outZip, [IO.FileMode]::Create)
    $archive = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($f in Get-ChildItem -LiteralPath $zipWork -Recurse -File) {
            $rel = $f.FullName.Substring($baseLen).Replace('\', '/')
            $entry = $archive.CreateEntry($rel, [IO.Compression.CompressionLevel]::Optimal)
            $es = $entry.Open()
            $bytes = [IO.File]::ReadAllBytes($f.FullName)
            $es.Write($bytes, 0, $bytes.Length)
            $es.Close()
        }
    }
    finally {
        $archive.Dispose()
        $fs.Close()
    }
    Move-Item -LiteralPath $outZip -Destination $xlamPath -Force
    Remove-Item -LiteralPath $tmpZip -Force -ErrorAction SilentlyContinue
    Write-Host '已注入 Ribbon XML'
}

Inject-Ribbon $outXlam $ribbonXml

# ---------- 4. 校验 zip 内容（条目名必须全部是正斜杠，否则 Excel 读不到 Ribbon） ----------
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
$zip = [IO.Compression.ZipFile]::OpenRead($outXlam)
$bad = @($zip.Entries | Where-Object { $_.FullName -like '*\*' })
$entryNames = @($zip.Entries | ForEach-Object { $_.FullName })
$entryCount = $zip.Entries.Count
$zip.Dispose()

foreach ($n in $entryNames) { Write-Host ("  zip 条目: " + $n) }
if ($bad.Count -gt 0) {
    Write-Error ("打包错误：以下条目名含反斜杠，Excel 将无法解析 → " + ($bad | ForEach-Object { $_.FullName }) -join ', ')
}
if (-not ($entryNames -contains 'customui/customui.xml')) {
    Write-Error '打包错误：缺少 customui/customui.xml 部件'
}
$relsCheck = $entryNames -contains '_rels/.rels'
$zipCheck = [IO.Compression.ZipFile]::OpenRead($outXlam)
$sr = New-Object IO.StreamReader(($zipCheck.Entries | Where-Object { $_.FullName -eq '_rels/.rels' }).Open())
$relsText = $sr.ReadToEnd(); $sr.Close(); $zipCheck.Dispose()
if ($relsText -notmatch 'extensibility') {
    Write-Error '打包错误：_rels/.rels 中缺少 ui/extensibility 关系，Excel 不会加载 Ribbon'
}
if ($relsText -notmatch 'customui/customui.xml') {
    Write-Error '打包错误：_rels/.rels 中的关系目标不是 customui/customui.xml'
}
Write-Host '  Ribbon 接线校验通过（_rels/.rels → customui/customui.xml）'
Write-Host ("  zip 条目总数: " + $entryCount + "（条目名均为正斜杠）")
Write-Host '== 构建完成 =='
